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
#   SUR    -> lógica REAL conectada (scripts de la Etapa 1)
#   NORTE  -> solo lectura (pendiente de conectar)
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

# Archivos que consolidar.py necesita y que control_calidad.py no busca.
# Cada zona tiene los suyos: Sur suma la base de proveedores, que su
# control de calidad no mira pero su consolidación sí.
NOMBRES_EXTRA_ZONA = {
    "RM": {
        "destinatarios": NOMBRES_MAESTROS["destinatarios"],
        "homologacion": NOMBRES_MAESTROS["homologacion"],
    },
    "SUR": {
        "destinatarios": NOMBRES_MAESTROS["destinatarios"],
        "homologacion": NOMBRES_MAESTROS["homologacion"],
        "proveedores": ["BBDD PROVEEDORES MOWI.xlsx", "BBDD PROVEEDORES.xlsx"],
    },
}

# 'destinatarios' es obligatorio para consolidar; 'homologacion' no lo es
# (el script original avisa y sigue sin aplicar correcciones). En Sur,
# 'proveedores' también es opcional: consolidar_sur lo omite si no está.
EXTRA_OBLIGATORIOS_ZONA = {"RM": ["destinatarios"], "SUR": ["destinatarios"]}

# Nombre base del consolidado que genera cada zona en modo prueba
NOMBRE_CONSOLIDADO = {"RM": "TRAZABILIDAD_RM.xlsx", "SUR": "TRAZABILIDAD_SUR.xlsx"}

# Zonas con la lógica real conectada. Las que no están acá entran en modo
# solo lectura: se comprueba que los archivos se puedan abrir y nada más.
ZONAS_CONECTADAS = ("RM", "SUR")

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

# Sur devuelve otras claves: su control de calidad no está numerado C1–C9.
ETIQUETAS_CC_SUR = {
    "c1": "Cliente con mismo nombre y RUT distinto",
    "c2": "Cliente con mismo RUT y nombre distinto",
    "c3": "Generador con variantes de nombre",
    "c_tipo": "TIPO sin código SINADER",
    "c_trans": "Transportistas sin RUT",
    "c_vacios": "Vacíos críticos",
    "c_mov": "Movimientos no reconocidos (se descartan al consolidar)",
    "c_dest": "Destinos vacíos sin explicación (se descartan al consolidar)",
}

ETIQUETAS_CC_ZONA = {"RM": ETIQUETAS_CC, "SUR": ETIQUETAS_CC_SUR}

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
                    # Se copia por trozos: un solo archivo del ZIP puede pesar
                    # decenas de MB y leerlo entero de golpe duplica la memoria.
                    with z.open(info) as origen, open(destino, "wb") as salida:
                        while True:
                            trozo = origen.read(1 << 20)
                            if not trozo:
                                break
                            salida.write(trozo)
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


def _dimension(z, ruta_hoja):
    """Rango que ocupa una hoja, leyendo solo el encabezado de su XML."""
    import xml.etree.ElementTree as ET

    with z.open(ruta_hoja) as fh:
        for ev, el in ET.iterparse(fh, events=("start",)):
            if el.tag == _NS + "dimension":
                return el.get("ref")
            if el.tag == _NS + "sheetData":
                return None
    return None


def _inventario(rutas):
    """Arma la tabla 'Archivo / Hojas / Filas' sin abrir los libros.

    Se leen los metadatos del .xlsx (que por dentro es un .zip con XML): los
    nombres de las hojas y el rango declarado de la primera. No se carga ni una
    celda, así que da lo mismo si el archivo pesa 60 MB.
    """
    filas = []
    for ruta in rutas:
        nombre = os.path.basename(ruta)
        try:
            with zipfile.ZipFile(ruta) as z:
                hojas, nombres_hojas = _hojas_del_libro(z)
                total = "—"
                if hojas:
                    ref = _dimension(z, hojas[0])
                    if ref and ":" in ref:
                        # se descuenta la fila de encabezado, para que la cifra
                        # sea comparable con la del resto del sistema
                        total = max(0, _partes_ref(ref.split(":")[1])[1] - 1)
            filas.append({"archivo": nombre,
                          "hojas": ", ".join(nombres_hojas),
                          "filas": total})
        except Exception as e:
            filas.append({"archivo": nombre,
                          "hojas": "no se pudo leer",
                          "filas": f"{type(e).__name__}: {e}"[:200]})
    return filas


# =============================================================================
# RECORTE MENSUAL
# -----------------------------------------------------------------------------
# El consolidado que producen los scripts NO es "el mes": es todo el año
# rehecho desde cero (el único filtro de tiempo es AÑO_DESDE). Eso no se toca,
# porque es exactamente lo que quedó verificado en la Etapa 1.
#
# Pero la operadora no pega el año entero: pega SOLO las filas del mes en la
# base maestra. Así que además del consolidado completo, acá se genera un
# archivo aparte con el recorte del mes elegido, listo para copiar y pegar.
#
# El recorte se hace por la columna 'Fecha' (que es una fecha de verdad), NO
# por la columna 'Mes': esa última guarda el nombre del mes sin el año
# ("agosto"), así que no distingue agosto de 2026 de agosto de 2027.
# =============================================================================
_MESES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
    7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE",
    12: "DICIEMBRE",
}


def _leer_periodo(periodo):
    """Convierte '2026-08' en (2026, 8). Devuelve None si no vino nada válido."""
    if not periodo:
        return None
    m = re.match(r"^\s*(\d{4})-(\d{1,2})\s*$", str(periodo))
    if not m:
        return None
    anio, mes = int(m.group(1)), int(m.group(2))
    if not 1 <= mes <= 12:
        return None
    return anio, mes


def _etiqueta_periodo(anio, mes):
    return f"{_MESES[mes]} {anio}"


def _recorte_mensual(df, periodo, carpeta_salida, log):
    """Genera el archivo con solo las filas del mes elegido.

    Devuelve {archivo, filas, kilos, alertas, resumen_meses} o None si no se
    pidió período o el consolidado no trae columna 'Fecha'.
    """
    par = _leer_periodo(periodo)
    if par is None:
        log.append("  (no se eligió período: no se genera el recorte mensual)")
        return None
    anio, mes = par
    etiqueta = _etiqueta_periodo(anio, mes)

    if "Fecha" not in df.columns:
        log.append("  ⚠ El consolidado no tiene columna 'Fecha': no se puede recortar.")
        return None

    fechas = pd.to_datetime(df["Fecha"], errors="coerce")
    sin_fecha = int(fechas.isna().sum())

    del_mes = df[(fechas.dt.year == anio) & (fechas.dt.month == mes)]
    posteriores = df[(fechas.dt.year > anio) |
                     ((fechas.dt.year == anio) & (fechas.dt.month > mes))]

    # Resumen de todo el consolidado, mes por mes. Es la herramienta para
    # detectar filas atrasadas: si un mes ya pegado creció respecto de la
    # corrida anterior, esas filas nuevas son cargas registradas tarde.
    col_peso = next((c for c in df.columns if "PESO" in str(c).upper()), None)
    resumen = pd.DataFrame({"MES": fechas.dt.to_period("M").astype(str)})
    resumen["PESO"] = (
        pd.to_numeric(df[col_peso], errors="coerce").fillna(0).values
        if col_peso else 0
    )
    resumen_meses = (
        resumen[resumen["MES"] != "NaT"]
        .groupby("MES")
        .agg(FILAS=("MES", "size"), PESO_TOTAL_KG=("PESO", "sum"))
        .reset_index()
        .sort_values("MES")
    )

    nombre = f"MES_{anio}-{mes:02d}_RM.xlsx"
    ruta = Path(carpeta_salida) / nombre
    hoja_mes = f"{_MESES[mes]}_{anio}"[:31]
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        del_mes.to_excel(w, sheet_name=hoja_mes, index=False)

    log.append(f"  Recorte de {etiqueta}: {len(del_mes)} fila(s) → {nombre}")

    # El desglose por mes va al registro de abajo, no al Excel: sirve para
    # comparar con la corrida anterior y detectar cargas registradas tarde,
    # pero no es algo que la operadora tenga que pegar en ninguna parte.
    log.append("  Desglose del consolidado completo:")
    for _, fila in resumen_meses.iterrows():
        log.append(f"    {fila['MES']}  {int(fila['FILAS']):>7} fila(s)"
                   f"  {_formato_miles(float(fila['PESO_TOTAL_KG']))} kg")

    # ---- Avisos ------------------------------------------------------------
    alertas = []
    if len(del_mes) == 0:
        alertas.append({
            "titulo": f"No hay ninguna fila de {etiqueta}",
            "detalle": "El consolidado no contiene datos de ese mes. Revisa si "
                       "la carpeta que subiste es la del mes correcto, o si "
                       "elegiste bien el período.",
        })
    if len(posteriores) > 0:
        meses_post = sorted(set(
            pd.to_datetime(posteriores["Fecha"], errors="coerce")
            .dt.to_period("M").astype(str)
        ))
        alertas.append({
            "titulo": f"Hay {len(posteriores)} fila(s) con fecha posterior a {etiqueta}",
            "detalle": "Meses encontrados: " + ", ".join(meses_post) +
                       ". Esas filas NO están en el archivo del mes. Puede ser "
                       "un error de tipeo en la fecha, o que subiste una "
                       "carpeta más nueva de la que creías.",
        })
    if sin_fecha > 0:
        alertas.append({
            "titulo": f"Hay {sin_fecha} fila(s) sin fecha válida",
            "detalle": "No se pueden asignar a ningún mes, así que quedaron "
                       "fuera del archivo del mes. Aparecen en el control de "
                       "calidad como campo crítico vacío.",
        })

    return {
        "archivo": nombre,
        "etiqueta_periodo": etiqueta,
        "filas": len(del_mes),
        "kilos": _sumar_kilos(del_mes),
        "alertas": alertas,
        "meses": len(resumen_meses),
    }


# =============================================================================
# REVISIONES ACOTADAS AL MES
# -----------------------------------------------------------------------------
# El control de calidad y la revisión del consolidado miran SIEMPRE el año
# completo (el filtro de los scripts es AÑO_DESDE, nunca un mes). Eso significa
# que un RUT mal escrito en marzo vuelve a aparecer en julio, en agosto y en
# septiembre, hasta que alguien lo corrija en el origen.
#
# Para que la operadora revise lo del mes que está cerrando, acá se recortan
# las hojas que tienen fecha, dejando solo las filas de ese mes. No se toca ni
# un cálculo: los scripts corren igual y sobre el año completo. Esto pasa
# DESPUÉS, sobre los Excel ya generados.
#
# Lo que NO se recorta:
#   · RESUMEN e INFO — son las cabeceras del archivo
#   · las hojas que ya son un resumen por mes (DESTINOS_POR_MES, V6, …)
#   · las hojas sin fecha (conflictos de RUT, variantes de nombre, columnas):
#     no son de un mes, son del catálogo completo
#
# Y para que nada quede escondido, se agrega una hoja RESUMEN_DEL_MES que dice,
# hoja por hoja, cuántos casos son del mes y cuántos hay en todo el año.
# =============================================================================
_HOJAS_SIN_RECORTE = ("RESUMEN", "INFO")


def _columna_fecha(df):
    """Encuentra la columna de fecha de una hoja, si la tiene."""
    for col in df.columns:
        # _clave() deja el nombre en mayúsculas y sin espacios ni guiones bajos:
        # "Fec_salida" queda "FECSALIDA" y "Fecha" queda "FECHA".
        limpio = _clave(str(col))
        if limpio.startswith("FECHA") or limpio == "FECSALIDA":
            return col
    return None


def _recortar_revision(ruta, anio, mes, etiqueta, log):
    """Deja en el Excel solo las filas del mes, en las hojas que tienen fecha.

    Devuelve el nombre del archivo nuevo, o None si no se pudo procesar.
    """
    ruta = Path(ruta)
    if not ruta.exists():
        return None

    try:
        hojas = pd.read_excel(ruta, sheet_name=None)
    except Exception as e:
        log.append(f"  ⚠ No se pudo abrir {ruta.name} para recortar: {e}")
        return None

    resumen, salida = [], {}
    for nombre, df in hojas.items():
        arriba = nombre.upper()
        col = _columna_fecha(df)
        recortable = (
            col is not None
            and not df.empty
            and arriba not in _HOJAS_SIN_RECORTE
            and not arriba.endswith("POR_MES")
            and "MENSUAL" not in arriba
        )

        if not recortable:
            salida[nombre] = df
            resumen.append({
                "HOJA": nombre,
                "CASOS DEL MES": "—",
                "CASOS EN EL AÑO": len(df),
                "RECORTADA": "no",
            })
            continue

        fechas = pd.to_datetime(df[col], errors="coerce")
        del_mes = df[(fechas.dt.year == anio) & (fechas.dt.month == mes)].copy()
        del_mes.insert(0, "MES", f"{anio}-{mes:02d}")
        salida[nombre] = del_mes
        resumen.append({
            "HOJA": nombre,
            "CASOS DEL MES": len(del_mes),
            "CASOS EN EL AÑO": len(df),
            "RECORTADA": "sí",
        })

    nuevo = ruta.with_name(f"{ruta.stem}_{anio}-{mes:02d}{ruta.suffix}")
    with pd.ExcelWriter(nuevo, engine="openpyxl") as w:
        pd.DataFrame({
            "Campo": ["Mes revisado", "Alcance de los cálculos"],
            "Valor": [etiqueta,
                      "Los scripts revisan el año completo. Las hojas con "
                      "fecha se recortaron a este mes; el total del año está "
                      "en RESUMEN_DEL_MES."],
        }).to_excel(w, sheet_name="MES_REVISADO", index=False)
        pd.DataFrame(resumen).to_excel(w, sheet_name="RESUMEN_DEL_MES", index=False)
        for nombre, df in salida.items():
            df.to_excel(w, sheet_name=nombre[:31], index=False)

    recortadas = sum(1 for r in resumen if r["RECORTADA"] == "sí")
    log.append(f"  {ruta.name} → {nuevo.name} ({recortadas} hoja(s) recortada(s) a {etiqueta})")
    return nuevo.name


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
# LECTURA LIVIANA DE TABLAS CON NOMBRE
# -----------------------------------------------------------------------------
# Los scripts leen las tablas con load_workbook(path), que carga el libro ENTERO
# en memoria creando un objeto por cada celda. En un computador con 16 GB eso da
# lo mismo; dentro del navegador el motor tiene un techo mucho más bajo y
# revienta con MemoryError.
#
# El caso que lo destapó: BBDD TRADICIONALES.xlsx pesa 61 MB comprimido y 450 MB
# descomprimido, pero el 97% de ese peso está en una hoja ("Base", 234.104 filas)
# que los scripts NUNCA leen. La tabla que sí se necesita ("Analizado") tiene
# 4.171 filas. Es decir: el navegador se moría cargando datos que no hacen falta.
#
# Estas funciones leen el .xlsx como lo que realmente es —un .zip con XML
# adentro— y recorren SOLO la hoja donde vive la tabla pedida, fila por fila,
# deteniéndose al llegar al final del rango. Nunca abren las hojas grandes.
#
# El resultado es idéntico al de openpyxl (verificado celda por celda sobre el
# archivo real y sobre un libro de prueba con fechas, horas, booleanos, decimales
# y celdas vacías). Los scripts de la Etapa 1 no se tocan: se les reemplaza el
# lector en el momento de usarlos.
# =============================================================================
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NSR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _hojas_del_libro(z):
    """Devuelve la lista de rutas internas de las hojas, en el orden del libro."""
    import xml.etree.ElementTree as ET

    rels = {}
    if "xl/_rels/workbook.xml.rels" in z.namelist():
        for rel in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")):
            rels[rel.get("Id")] = rel.get("Target", "").lstrip("/")

    hojas, nombres = [], []
    libro = ET.fromstring(z.read("xl/workbook.xml"))
    for hoja in libro.find(_NS + "sheets"):
        destino = rels.get(hoja.get(_NSR + "id"), "")
        if destino:
            hojas.append(destino if destino.startswith("xl/") else "xl/" + destino)
            nombres.append(hoja.get("name"))
    return hojas, nombres


def _hojas_y_tablas(ruta):
    """Devuelve {nombre_de_tabla: (ruta_interna_de_la_hoja, rango)}."""
    import xml.etree.ElementTree as ET

    resultado = {}
    with zipfile.ZipFile(ruta) as z:
        nombres = set(z.namelist())
        hojas, _ = _hojas_del_libro(z)

        # Definición de cada tabla: nombre y rango
        tablas = {}
        for n in nombres:
            if n.startswith("xl/tables/") and n.endswith(".xml"):
                t = ET.fromstring(z.read(n))
                tablas[n.split("/")[-1]] = (
                    t.get("name") or t.get("displayName"), t.get("ref")
                )

        # Qué tabla pertenece a qué hoja
        for hoja in hojas:
            archivo = hoja.split("/")[-1]
            rels = f"xl/worksheets/_rels/{archivo}.rels"
            if rels not in nombres:
                continue
            for rel in ET.fromstring(z.read(rels)):
                destino = rel.get("Target", "")
                if "tables/" in destino:
                    nombre, ref = tablas.get(destino.split("/")[-1], (None, None))
                    if nombre and ref:
                        resultado[nombre] = (hoja, ref)

    return resultado


def _col_a_numero(letras):
    n = 0
    for ch in letras:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _partes_ref(ref):
    m = re.match(r"([A-Za-z]+)(\d+)", ref)
    return _col_a_numero(m.group(1)), int(m.group(2))


def _estilos_de_fecha(z):
    """Qué estilos de celda representan fechas/horas (para convertir el número)."""
    import xml.etree.ElementTree as ET
    from openpyxl.styles.numbers import is_date_format, BUILTIN_FORMATS

    if "xl/styles.xml" not in z.namelist():
        return set(), set()

    raiz = ET.fromstring(z.read("xl/styles.xml"))
    propios = {}
    fmts = raiz.find(_NS + "numFmts")
    if fmts is not None:
        for f in fmts:
            propios[int(f.get("numFmtId"))] = f.get("formatCode")

    fechas, duraciones = set(), set()
    xfs = raiz.find(_NS + "cellXfs")
    if xfs is None:
        return fechas, duraciones
    for i, xf in enumerate(xfs):
        nid = int(xf.get("numFmtId", 0))
        codigo = propios.get(nid) or BUILTIN_FORMATS.get(nid)
        if codigo and is_date_format(codigo):
            fechas.add(i)
            if re.search(r"\[(h+|m+|s+)\]", codigo):
                duraciones.add(i)
    return fechas, duraciones


def _a_numero(txt):
    if "." in txt or "E" in txt or "e" in txt:
        return float(txt)
    try:
        return int(txt)
    except ValueError:
        return float(txt)


def _leer_rango(ruta, ruta_hoja, ref):
    """Lee un rango recorriendo el XML de la hoja. Devuelve lista de tuplas."""
    import xml.etree.ElementTree as ET
    from openpyxl.utils.datetime import from_excel, CALENDAR_WINDOWS_1900

    ini, fin = ref.split(":") if ":" in ref else (ref, ref)
    c_min, f_min = _partes_ref(ini)
    c_max, f_max = _partes_ref(fin)
    ancho = c_max - c_min + 1

    with zipfile.ZipFile(ruta) as z:
        fechas, duraciones = _estilos_de_fecha(z)

        # --- Pasada 1: la hoja. Los textos quedan anotados como pendientes.
        filas, pendientes = {}, set()
        with z.open(ruta_hoja) as fh:
            n_fila = 0
            for _, el in ET.iterparse(fh, events=("end",)):
                if el.tag != _NS + "row":
                    if el.tag == _NS + "sheetData":
                        el.clear()
                    continue
                n_fila = int(el.get("r") or (n_fila + 1))
                if f_min <= n_fila <= f_max:
                    vals = [None] * ancho
                    for c in el:
                        if c.tag != _NS + "c":
                            continue
                        r = c.get("r")
                        if not r:
                            continue
                        ci = _partes_ref(r)[0]
                        if not (c_min <= ci <= c_max):
                            continue
                        t = c.get("t")
                        if t == "inlineStr":
                            bloque = c.find(_NS + "is")
                            v = ("".join(x.text or "" for x in bloque.iter(_NS + "t"))
                                 if bloque is not None else None)
                        else:
                            nodo = c.find(_NS + "v")
                            v = nodo.text if nodo is not None else None
                            if v is None or t in ("str", "e"):
                                pass
                            elif t == "s":
                                idx = int(v)
                                pendientes.add(idx)
                                v = ("\x00SST", idx)
                            elif t == "b":
                                v = v == "1"
                            else:
                                v = _a_numero(v)
                                sid = int(c.get("s") or 0)
                                if sid in fechas:
                                    v = from_excel(v, CALENDAR_WINDOWS_1900,
                                                   timedelta=sid in duraciones)
                        vals[ci - c_min] = v
                    filas[n_fila] = vals
                el.clear()
                if n_fila > f_max:
                    break

        # --- Pasada 2: SOLO los textos que el rango realmente usa.
        sst = {}
        if pendientes and "xl/sharedStrings.xml" in z.namelist():
            with z.open("xl/sharedStrings.xml") as fh:
                i = 0
                for _, el in ET.iterparse(fh, events=("end",)):
                    if el.tag != _NS + "si":
                        continue
                    if i in pendientes:
                        sst[i] = "".join(x.text or "" for x in el.iter(_NS + "t"))
                    i += 1
                    el.clear()

    salida = []
    for n in range(f_min, f_max + 1):
        vals = filas.get(n)
        if vals is None:
            salida.append(tuple([None] * ancho))
        else:
            salida.append(tuple(
                sst.get(v[1]) if type(v) is tuple and v[0] == "\x00SST" else v
                for v in vals
            ))
    return salida


def _tabla_a_dataframe(ruta, nombre_tabla, p, limpiar_encabezados):
    ruta = Path(ruta)
    p(f"  Leyendo {ruta.name} → tabla {nombre_tabla}...")

    tablas = _hojas_y_tablas(ruta)
    if nombre_tabla not in tablas:
        raise ValueError(
            f"El archivo {ruta.name} no tiene una tabla llamada '{nombre_tabla}'.\n"
            f"  Revisa que la tabla no haya sido renombrada o eliminada en Excel."
        )

    hoja, ref = tablas[nombre_tabla]
    filas = _leer_rango(ruta, hoja, ref)
    if not filas:
        return pd.DataFrame()

    df = pd.DataFrame(list(filas[1:]), columns=list(filas[0]))
    if limpiar_encabezados:
        df.columns = df.columns.astype(str).str.strip()
    return df.dropna(how="all").reset_index(drop=True)


def _primera_tabla_de_hoja(ruta, nombre_hoja):
    """Devuelve ((ruta_interna, rango), titulos) de la primera tabla de una hoja.

    Es el equivalente liviano de abrir el libro y mirar ws.tables. Si la hoja
    existe pero no tiene ninguna tabla, la primera parte viene en None.
    """
    import xml.etree.ElementTree as ET

    with zipfile.ZipFile(ruta) as z:
        nombres = set(z.namelist())
        hojas, titulos = _hojas_del_libro(z)
        if nombre_hoja not in titulos:
            return None, titulos

        hoja = hojas[titulos.index(nombre_hoja)]
        rels = f"xl/worksheets/_rels/{hoja.split('/')[-1]}.rels"
        if rels not in nombres:
            return None, titulos

        for rel in ET.fromstring(z.read(rels)):
            destino = rel.get("Target", "")
            if "tables/" in destino:
                parte = "xl/tables/" + destino.split("/")[-1]
                if parte in nombres:
                    ref = ET.fromstring(z.read(parte)).get("ref")
                    if ref:
                        return (hoja, ref), titulos
    return None, titulos


def _instalar_lectores_livianos(mod_control, mod_consolidar, log=None):
    """Reemplaza los lectores por versiones que no agotan la memoria.

    Hay dos formatos de lector, uno por zona:

      · RM lee por NOMBRE DE TABLA        → leer_tabla / leer_tabla_nombrada
      · Sur lee por NOMBRE DE HOJA        → leer_fuente

    Cada uno conserva el post-proceso de su script original: control_calidad de
    RM limpia los espacios de los encabezados y consolidar no, y eso NO se
    cambia. En Sur los dos limpian, igual que el original.

    Si una hoja de Sur no tiene tabla Excel, se deja correr el lector original:
    ahí el script adivina la fila de encabezado probando varias, y esa lógica
    no se toca.
    """
    def leer_tabla(path, nombre_tabla, p):
        return _tabla_a_dataframe(path, nombre_tabla, p, limpiar_encabezados=True)

    def leer_tabla_nombrada(path, nombre_tabla, p):
        return _tabla_a_dataframe(path, nombre_tabla, p, limpiar_encabezados=False)

    if hasattr(mod_control, "leer_tabla"):
        mod_control.leer_tabla = leer_tabla
    if hasattr(mod_consolidar, "leer_tabla_nombrada"):
        mod_consolidar.leer_tabla_nombrada = leer_tabla_nombrada

    def hacer_leer_fuente(original):
        def leer_fuente(path, hoja, p, headers_fallback=(0, 7, 8, 9)):
            ruta = Path(path)
            encontrada, titulos = _primera_tabla_de_hoja(ruta, hoja)

            if hoja not in titulos:
                raise ValueError(
                    f"No existe la hoja '{hoja}' en {ruta.name}. Hojas: {titulos}"
                )
            if encontrada is None:
                # Sin tabla: lo resuelve el lector original, que prueba varias
                # filas de encabezado y se queda con la que más columnas
                # conocidas encuentra.
                return original(path, hoja, p, headers_fallback=headers_fallback)

            p(f"  Leyendo {ruta.name} → hoja {hoja}...")
            ruta_hoja, ref = encontrada
            filas = _leer_rango(ruta, ruta_hoja, ref)
            if not filas:
                return pd.DataFrame()

            df = pd.DataFrame(list(filas[1:]), columns=list(filas[0]))
            df.columns = df.columns.astype(str).str.strip()
            return df.dropna(how="all").reset_index(drop=True)
        return leer_fuente

    for modulo in (mod_control, mod_consolidar):
        if hasattr(modulo, "leer_fuente"):
            modulo.leer_fuente = hacer_leer_fuente(modulo.leer_fuente)

    if log is not None:
        log.append("  (lectura de tablas en modo liviano, para no agotar la memoria)")


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
        # Lista completa de los cuatro maestros. La app la necesita para saber
        # qué falta mirando TODO lo guardado, no solo lo de este envío.
        "esperados": [
            {"clave": k, "etiqueta": ETIQUETAS_MAESTROS[k], "nombre": v[0]}
            for k, v in NOMBRES_MAESTROS.items()
        ],
        "log": "\n".join(log),
    }


# =============================================================================
# ZONA RM — LÓGICA REAL
# =============================================================================
def _rutas_zona(zona, carpeta_entrada, control_calidad):
    """Arma el diccionario de rutas que necesitan los scripts de una zona.

    Parte de rutas_desde_carpeta() del control de calidad de esa zona (que ya
    encuentra las fuentes tolerando tildes y mayúsculas) y le agrega los
    archivos que solo necesita consolidar().
    """
    NOMBRES_EXTRA = NOMBRES_EXTRA_ZONA.get(zona, {})
    EXTRA_OBLIGATORIOS = EXTRA_OBLIGATORIOS_ZONA.get(zona, [])

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


def _alertas_zona(zona, cc, rev, periodo_txt=None):
    """Traduce los resultados de los scripts a la lista de alertas de pantalla.

    Las cifras son SIEMPRE del año completo, porque así corren los scripts. Si
    hay un mes elegido, se dice explícitamente, para que nadie confunda estos
    números con los del archivo recortado.
    """
    alertas = []
    if periodo_txt:
        alertas.append({
            "titulo": f"Las cifras de abajo son del año completo, no de {periodo_txt}",
            "detalle": "Los scripts revisan todo 2026 en cada corrida, así que "
                       "los errores de meses anteriores siguen apareciendo hasta "
                       "que se corrijan en el origen. Cuántos son de "
                       f"{periodo_txt} está en la hoja RESUMEN_DEL_MES de cada "
                       "archivo de revisión.",
        })

    for clave, etiqueta in ETIQUETAS_CC_ZONA.get(zona, ETIQUETAS_CC).items():
        df = cc.get(clave)
        if df is not None and len(df) > 0:
            alertas.append({
                "titulo": etiqueta,
                "detalle": f"{len(df)} caso(s) en el año. El detalle está en el "
                           f"archivo de control de calidad.",
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


def _procesar_zona(zona, carpeta_entrada, carpeta_salida, log, periodo=None):
    """Ejecuta el flujo completo de una zona: control de calidad → consolidar → revisar.

    Sirve igual para RM y para Sur. Lo único que cambia entre zonas está en las
    tablas de configuración de arriba (qué archivos extra pide la consolidación,
    cómo se llaman los controles) y en la carpeta de la que se importan los
    scripts: python/rm/ o python/sur/.
    """
    import gc
    import importlib

    paquete = zona.lower()
    mod_consolidar = importlib.import_module(f"{paquete}.consolidar")
    mod_control = importlib.import_module(f"{paquete}.control_calidad")
    mod_revisar = importlib.import_module(f"{paquete}.revisar_consolidado")

    _instalar_lectores_livianos(mod_control, mod_consolidar, log)

    salida = Path(carpeta_salida)
    nombre_cc = f"CONTROL_CALIDAD_{zona}.xlsx"
    nombre_rev = f"REVISION_CONSOLIDADO_{zona}.xlsx"

    # ---- 1. Ubicar los archivos --------------------------------------------
    log.append("── Buscando los archivos ──")
    rutas = _rutas_zona(zona, carpeta_entrada, mod_control)
    for clave in sorted(rutas):
        if clave != "destino_real":
            log.append(f"  {clave:<15} → {Path(rutas[clave]).name}")
    log.append("")

    # ---- 2. Control de calidad (C1 a C9) ------------------------------------
    #
    # OJO: control_calidad.verificar_rutas() comprueba que exista TODA clave
    # del diccionario que recibe, no una lista fija. Si se le pasa el
    # diccionario completo de consolidar(), reclama por 'destino_real', que a
    # propósito apunta a un archivo inexistente. Por eso se le entregan solo
    # las claves que ese script declara como suyas.
    claves_cc = set(getattr(mod_control, "NOMBRES_ARCHIVO", {}))
    rutas_cc = {k: v for k, v in rutas.items() if k in claves_cc} or rutas

    log.append("── Control de calidad ──")
    cc = mod_control.controlar(
        rutas_cc,
        ruta_salida=salida / nombre_cc,
        mostrar=True,   # el avance se ve en pantalla mientras corre
    )
    log.append(cc.get("log", ""))
    log.append("")

    # Se sueltan los DataFrames del control de calidad antes de consolidar:
    # dentro del navegador la memoria es el recurso escaso.
    for clave in list(cc):
        if clave.startswith(("c", "v")) and clave not in ("total",):
            pass
    gc.collect()

    # ---- 3. Consolidación (modo prueba: no toca la base real) ---------------
    log.append("── Consolidación ──")
    antes = {f.name for f in salida.glob("*.xlsx")}
    res = mod_consolidar.consolidar(
        rutas,
        ruta_prueba=salida / NOMBRE_CONSOLIDADO.get(zona, f"TRAZABILIDAD_{zona}.xlsx"),
        ruta_log=None,
        modo_reset=False,
        modo_prueba=True,     # nunca escribe sobre BBDD_TRAZABILIDAD_RM.xlsx
        mostrar=True,
    )
    log.append(res.get("log", ""))

    nombre_consolidado = _archivo_nuevo(salida, antes)
    if not nombre_consolidado:
        raise RuntimeError(
            "La consolidación terminó pero no se encontró el archivo generado."
        )

    # consolidar() le pone al archivo la fecha y hora de la corrida. Eso dice
    # CUÁNDO se generó, no QUÉ período cubre. Si la operadora eligió un mes, el
    # archivo se renombra para que se entienda solo dentro de tres meses.
    par = _leer_periodo(periodo)
    if par:
        nuevo = f"TRAZABILIDAD_{zona}_{par[0]}-{par[1]:02d}.xlsx"
        try:
            (salida / nombre_consolidado).replace(salida / nuevo)
            log.append(f"  Archivo renombrado a {nuevo}")
            nombre_consolidado = nuevo
        except OSError as e:
            log.append(f"  (no se pudo renombrar el consolidado: {e})")
    log.append("")

    # ---- 4. Revisión del consolidado (V0 a V9) ------------------------------
    log.append("── Revisión del consolidado ──")
    rev = mod_revisar.revisar(
        salida / nombre_consolidado,
        ruta_salida=salida / nombre_rev,
        mostrar=True,
    )
    log.append(rev.get("log", ""))

    # ---- 5. Recorte del mes elegido -----------------------------------------
    log.append("")
    log.append("── Recorte del mes ──")
    mensual = _recorte_mensual(res["consolidado"], periodo, salida, log)

    # Las revisiones también se acotan al mes que está cerrando la operadora.
    cc_mes = rev_mes = None
    if par:
        etiqueta = _etiqueta_periodo(*par)
        cc_mes = _recortar_revision(
            salida / nombre_cc, par[0], par[1], etiqueta, log)
        rev_mes = _recortar_revision(
            salida / nombre_rev, par[0], par[1], etiqueta, log)

    # ---- 6. Armar lo que se muestra en pantalla -----------------------------
    kilos = _sumar_kilos(res["consolidado"])
    resumen = (
        f"<b>Zona {zona} consolidada.</b> "
        f"{_formato_miles(res['filas'], 0)} filas × {res['columnas']} columnas"
        + (f" · {_formato_miles(kilos)} kg" if kilos else "")
        + f" · control de calidad: {cc.get('total', 0)} conflicto(s)"
        + f" · revisión: {rev.get('total_alertas', 0)} alerta(s)."
    )
    if mensual:
        resumen += (
            f"<br><b>{mensual['etiqueta_periodo']}:</b> "
            f"{_formato_miles(mensual['filas'], 0)} fila(s) para pegar en la base"
            + (f" · {_formato_miles(mensual['kilos'])} kg" if mensual["kilos"] else "")
            + "."
        )

    descargas = []
    if mensual:
        descargas.append({
            "archivo": mensual["archivo"],
            "etiqueta": f"Mes {mensual['etiqueta_periodo']}",
        })
    descargas += [
        {"archivo": cc_mes or nombre_cc,
         "etiqueta": "Control de calidad" + (f" · {mensual['etiqueta_periodo']}"
                                             if cc_mes and mensual else "")},
        {"archivo": rev_mes or nombre_rev,
         "etiqueta": "Revisión del consolidado" + (f" · {mensual['etiqueta_periodo']}"
                                                   if rev_mes and mensual else "")},
        {"archivo": nombre_consolidado, "etiqueta": "Consolidado completo del año"},
    ]
    if cc_mes:
        descargas.append({"archivo": nombre_cc,
                          "etiqueta": "Control de calidad · año completo"})
    if rev_mes:
        descargas.append({"archivo": nombre_rev,
                          "etiqueta": "Revisión del consolidado · año completo"})

    alertas = (mensual["alertas"] if mensual else []) + _alertas_zona(
        zona, cc, rev, mensual["etiqueta_periodo"] if mensual else None)

    return {
        "resumen": resumen,
        "fuentes": _inventario(_listar_archivos(carpeta_entrada)),
        "alertas": alertas,
        "salida": mensual["archivo"] if mensual else nombre_consolidado,
        "descargas": descargas,
        "log": "\n".join(log),
    }


# =============================================================================
# ZONAS SUR Y NORTE — TODAVÍA EN DEMOSTRACIÓN
# =============================================================================
def _procesar_demostracion(zona, carpeta_entrada, carpeta_salida, log):
    """Revisa que los archivos de la zona se puedan leer. NO consolida nada.

    Se mantiene para SUR y NORTE mientras no se conecten sus scripts. Antes
    esta función pegaba las primeras hojas una debajo de otra y dejaba un
    CONSOLIDADO_<ZONA>.xlsx para descargar. Ese archivo parecía un resultado
    real y no lo era: no tenía control de calidad, ni homologación, ni filtros,
    ni recorte del mes. Si alguien lo pegaba en la base de trazabilidad, metía
    basura sin darse cuenta.

    Por eso ahora la zona en demostración solo INFORMA: dice qué archivos leyó,
    qué hojas tiene cada uno y cuántas filas trae la primera. No genera ningún
    Excel descargable.
    """
    rutas = _listar_archivos(carpeta_entrada)
    fuentes, alertas = [], []
    total_filas = 0

    for ruta in rutas:
        nombre = os.path.basename(ruta)
        try:
            with zipfile.ZipFile(ruta) as z:
                hojas_int, nombres_hojas = _hojas_del_libro(z)
                filas = 0
                if hojas_int:
                    ref = _dimension(z, hojas_int[0])
                    if ref and ":" in ref:
                        filas = max(0, _partes_ref(ref.split(":")[1])[1] - 1)
            fuentes.append({"archivo": nombre,
                            "hojas": ", ".join(nombres_hojas),
                            "filas": filas})
            total_filas += filas
            log.append(f"  {nombre}: hojas={nombres_hojas}, filas primera hoja={filas}")
        except Exception as e:
            fuentes.append({"archivo": nombre, "hojas": "no se pudo leer",
                            "filas": f"{type(e).__name__}: {e}"[:200]})
            alertas.append({"titulo": f"No se pudo leer {nombre}", "detalle": str(e)})
            log.append(f"  ERROR en {nombre}: {e}")

    ok = sum(1 for f in fuentes if f["hojas"] != "no se pudo leer")
    if fuentes:
        resumen = (
            f"<b>Zona {zona} — solo lectura.</b> {ok} de {len(rutas)} archivo(s) "
            f"se abrieron sin problema. <b>No se generó ningún consolidado</b>, "
            f"porque la lógica de esta zona todavía no está conectada."
        )
    else:
        resumen = "No se leyó ningún archivo válido."

    alertas.insert(0, {
        "titulo": f"Zona {zona}: todavía no está conectada",
        "detalle": "Por ahora la app solo comprueba que los archivos de esta "
                   "zona se puedan abrir. No hay control de calidad, ni "
                   "homologación, ni consolidación, ni recorte del mes: eso "
                   "llega cuando se conecten los scripts de la zona. "
                   "La única zona operativa es RM.",
    })

    log.append("")
    log.append(f"Lectura de prueba: {len(rutas)} archivo(s), {total_filas} fila(s) "
               f"en las primeras hojas. No se generó ningún archivo.")

    return {
        "resumen": resumen,
        "fuentes": fuentes,
        "alertas": alertas,
        "salida": None,
        "descargas": [],
        "log": "\n".join(log),
    }


# =============================================================================
# PUNTO DE ENTRADA — es lo que llama app.js
# =============================================================================
def procesar(zona, carpeta_entrada, carpeta_salida, periodo=None):
    zona = (zona or "").strip().upper()
    par = _leer_periodo(periodo)
    log = [
        f"Zona: {zona}",
        f"Período: {_etiqueta_periodo(*par) if par else '(no elegido)'}",
        f"Carpeta de entrada: {carpeta_entrada}",
        "",
    ]

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

        if zona in ZONAS_CONECTADAS:
            return _procesar_zona(zona, carpeta_entrada, carpeta_salida, log, periodo)
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
