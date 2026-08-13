"""
consolidar.py — Consolidador de Trazabilidad RM (versión refactorizada)
Ambipar Group — Proyecto TRZ-APP-001, Etapa 1

QUÉ CAMBIÓ respecto de 03_consolidar_rm.py:

    1. Las rutas ya no están escritas dentro del código: la función principal
       recibe un diccionario 'rutas' (igual que en control_calidad.py) más la
       ruta de HOMOLOGACION y la ruta de salida.

    2. El log de auditoría ya no se abre automáticamente al importar el
       archivo (antes lo hacía 'logging.basicConfig' apenas se cargaba el
       script). Ahora se abre solo cuando se llama a consolidar(), y solo si
       se le indica un archivo de log. Esto es necesario para que la app
       pueda importar este módulo sin que abra un archivo de inmediato.

    3. Cada función que antes leía RUTAS[...] del entorno global ahora la
       recibe como argumento. Esto es lo que permite consolidar RM tanto con
       los archivos de OneDrive como con archivos subidos por la operadora.

QUÉ NO CAMBIÓ:

    Todas las reglas de negocio son exactamente las mismas, línea por línea:
    los filtros, las exclusiones de plantas propias, los cálculos de
    movimiento interempresa, la inferencia de destino en Tradicionales, los
    joins con SINADER y Destinatarios, y el modo acumulativo. Nada de eso fue
    tocado.

    Se conserva también un detalle del código original que vale la pena
    señalar: la constante PLANTAS_PROPIAS_ACTIVAS está declarada pero no se
    usa en ningún cálculo — las exclusiones de GIRI y Ecofibras están
    escritas directamente en procesar_giri() y procesar_ecofibras(). Es así
    en el script original; se documenta acá para que quede visible, pero no
    se modifica sin que la dueña del proceso lo pida.

Se puede usar de dos formas:

    Desde otro programa:
        from consolidar import consolidar, rutas_desde_config
        rutas = rutas_desde_config()
        resultado = consolidar(rutas, modo_prueba=True)

    Desde la terminal, igual que antes:
        python consolidar.py --prueba
        python consolidar.py --reset
"""

import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook, Workbook


# ══════════════════════════════════════════════════════════════════
# CONSTANTES DE LA ZONA RM — reglas de negocio, no rutas
# ══════════════════════════════════════════════════════════════════
HOJA_DESTINO      = "TRAZABILIDAD"
AÑO_DESDE         = 2026
REGION_RM         = "Región Metropolitana"
REGION_VALPARAISO = "Región de Valparaíso"

# Hoja del archivo BBDD_DESTINATARIO.xlsx (antes se llamaba "RM Y ZONA SUR")
HOJA_DESTINATARIOS = "NACIONAL"

# Plantas propias activas en RM. Declarada por completitud y documentación:
# en el script original las exclusiones de traslados hacia estas plantas
# están escritas directamente en procesar_giri() y procesar_ecofibras(), no
# leen esta constante. Ver nota en el docstring de arriba.
PLANTAS_PROPIAS_ACTIVAS = {
    "ECOFIBRAS SAN BERNARDO",
    "GIRI",
}

COLUMNAS_FINALES = [
    "Fecha", "Mes", "Cliente", "RUT", "Gestor", "Contrato", "Generador",
    "Transportista", "Rut transportista", "Patente de Camión", "Ticket de pesaje",
    "Peso neto (kg)", "Destino", "Comuna Destino", "RUT DESTINATARIO",
    "TIPO", "CÓDIGOS SINADER", "Movimiento", "Movimiento interempresa",
    "CÓDIGO ESTABLECIMIENTO SINADER", "CÓDIGO DE TRATAMIENTO SINADER",
    "Región", "Destino inferido",
]

CLAVE_DEDUP = ["Fecha", "Cliente", "Ticket de pesaje", "TIPO", "Destino"]

MAX_REINTENTOS   = 5
ESPERA_REINTENTO = 30

MESES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
    7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}


class ArchivoFaltante(Exception):
    """Se lanza cuando falta algún archivo fuente necesario para consolidar."""


# ══════════════════════════════════════════════════════════════════
# REGISTRO DE MENSAJES — reemplaza al logging.basicConfig de nivel de módulo
# ══════════════════════════════════════════════════════════════════
class Registro:
    """Guarda los mensajes, los muestra, y opcionalmente los agrega a un log.

    El script original abría el archivo de log ni bien se importaba
    (logging.basicConfig se ejecutaba al cargar el módulo). Eso es un
    problema para una aplicación: importar el módulo no debería tener el
    efecto secundario de crear un archivo. Ahora el log solo se abre cuando
    se llama a consolidar() y se le indica dónde escribirlo, igual que hacía
    el script original (agregando líneas, sin borrar las anteriores).
    """

    def __init__(self, mostrar=True, ruta_log=None):
        self.lineas = []
        self.mostrar = mostrar
        self.ruta_log = Path(ruta_log) if ruta_log else None

    def __call__(self, msg=""):
        msg = str(msg)
        self.lineas.append(msg)
        if self.mostrar:
            print(msg)
        if self.ruta_log:
            marca = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(self.ruta_log, "a", encoding="utf-8") as f:
                f.write(f"{marca}  {msg}\n")

    def texto(self):
        return "\n".join(self.lineas)


# ══════════════════════════════════════════════════════════════════
# CÓMO ENCONTRAR LOS ARCHIVOS
# ══════════════════════════════════════════════════════════════════
def rutas_desde_config(base=None):
    """Arma las rutas de las fuentes y de las tablas auxiliares desde OneDrive.

    Devuelve un diccionario con dos partes:
      - las rutas de las 6 fuentes + destinatarios/sinader/transportistas
      - 'homologacion' y 'destino_real', que se usan aparte
    """
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from config import carpetas

    c = carpetas(base)
    zona = c["zona_centro"]
    info = c["info_base"]
    giri = c["giri"]

    return {
        "bo":             zona / "BBDD BO SAN BERNARDO.xlsx",
        "bo_valparaiso":  zona / "BBDD BO VALPARAÍSO.xlsx",
        "ecofibras":      zona / "BBDD ECOFIBRAS SAN BERNARDO.xlsx",
        "proveedores":    zona / "BBDD PROVEEDORES.xlsx",
        "tradicionales":  zona / "BBDD TRADICIONALES.xlsx",
        "giri":           giri / "OPERACIONES GIRI (INGRESOS Y EGRESOS).xlsx",
        "destinatarios":  info / "BBDD DESTINATARIO.xlsx",
        "sinader":        info / "Clasificación_Residuos SINADER.xlsx",
        "transportistas": info / "Transportistas.xlsx",
        "homologacion":   info / "HOMOLOGACION.xlsx",
        "destino_real":   c["bbdd"] / "BBDD_TRAZABILIDAD_RM.xlsx",
    }


def verificar_rutas(rutas):
    """Comprueba que las fuentes obligatorias existan antes de empezar a leer.

    HOMOLOGACION no es obligatoria: el script original sigue sin ella,
    avisando que continúa sin correcciones. destino_real tampoco, porque en
    modo prueba nunca se usa.
    """
    obligatorias = [
        "bo", "bo_valparaiso", "ecofibras", "proveedores", "tradicionales",
        "giri", "destinatarios", "sinader", "transportistas",
    ]
    faltan = [f"{k}: {rutas[k]}" for k in obligatorias if not Path(rutas[k]).exists()]
    if faltan:
        detalle = "\n".join(f"    · {f}" for f in faltan)
        raise ArchivoFaltante(f"No se encontraron estos archivos:\n{detalle}")


# ══════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════
def leer_excel(path, p, **kwargs):
    for intento in range(1, MAX_REINTENTOS + 1):
        try:
            return pd.read_excel(path, **kwargs)
        except PermissionError:
            p(f"  ⚠ Bloqueado: {path.name} — intento {intento}/{MAX_REINTENTOS}")
            if intento < MAX_REINTENTOS:
                time.sleep(ESPERA_REINTENTO)
    raise PermissionError(f"No se pudo abrir {path.name}")


def leer_tabla_nombrada(path, nombre_tabla, p):
    for intento in range(1, MAX_REINTENTOS + 1):
        try:
            wb = load_workbook(path, data_only=True)
            for sheet in wb.worksheets:
                for tabla in sheet.tables.values():
                    if tabla.name == nombre_tabla:
                        filas = [[c.value for c in fila] for fila in sheet[tabla.ref]]
                        headers = filas[0]
                        df = pd.DataFrame(filas[1:], columns=headers)
                        wb.close()
                        return df.dropna(how="all").reset_index(drop=True)
            wb.close()
            raise ValueError(f"Tabla '{nombre_tabla}' no encontrada en {path.name}")
        except PermissionError:
            p(f"  ⚠ Bloqueado: {path.name} — intento {intento}/{MAX_REINTENTOS}")
            if intento < MAX_REINTENTOS:
                time.sleep(ESPERA_REINTENTO)
    raise PermissionError(f"No se pudo abrir {path.name}")


def nombre_mes(fecha):
    try:
        return MESES_ES[pd.Timestamp(fecha).month]
    except Exception:
        return ""


def normalizar_nombre_para_buscar(texto):
    if pd.isna(texto):
        return ""
    return (
        str(texto).strip().upper()
        .replace(".", "")
        .replace("  ", " ")
        .strip()
    )


def limpiar_texto_serie(serie: pd.Series) -> pd.Series:
    return (
        serie.astype(str)
        .str.strip()
        .replace({"nan": None, "None": None, "": None})
    )


def limpiar_filas_invalidas(df: pd.DataFrame, p) -> pd.DataFrame:
    """Elimina filas completamente inválidas que vienen como totales o arrastres.

    Criterio conservador: solo elimina filas donde faltan simultáneamente
    Fecha, Cliente, Destino y TIPO. No elimina filas por patente, ticket o
    comuna vacía, porque esas deben mantenerse y alertarse.
    """
    requeridas = ["Fecha", "Cliente", "Destino", "TIPO"]
    existentes = [c for c in requeridas if c in df.columns]
    if not existentes:
        return df

    temp = df[existentes].copy()
    for c in existentes:
        temp[c] = temp[c].astype(str).str.strip().str.lower()
    vacia = temp.isin(["", "nan", "none", "nat"]).all(axis=1)

    n = int(vacia.sum())
    if n > 0:
        p(f"  ⚠ Filas inválidas eliminadas al final: {n} (sin Fecha, Cliente, Destino ni TIPO)")
    return df.loc[~vacia].copy()


def cargar_homologacion(ruta_homologacion, p):
    p("\n── Cargando tabla de homologación ──")

    homolog_clientes = {}
    homolog_generadores = {}

    if not Path(ruta_homologacion).exists():
        p("  ⚠ No se encontró HOMOLOGACION.xlsx — continuando sin correcciones")
        return homolog_clientes, homolog_generadores

    try:
        df_cli = leer_excel(ruta_homologacion, p, sheet_name="clientes")
        df_cli = df_cli.dropna(subset=["NOMBRE_VARIANTE"])
        for _, row in df_cli.iterrows():
            variante = normalizar_nombre_para_buscar(row["NOMBRE_VARIANTE"])
            nombre_correcto = str(row["NOMBRE_CORRECTO"]).strip() if pd.notna(row["NOMBRE_CORRECTO"]) else None
            rut_correcto = str(row["RUT_CORRECTO"]).strip() if pd.notna(row["RUT_CORRECTO"]) else None
            if variante and nombre_correcto:
                homolog_clientes[variante] = {
                    "nombre": nombre_correcto,
                    "rut": rut_correcto,
                }
        p(f"  Clientes homologados:    {len(homolog_clientes)} entradas")
    except Exception as e:
        p(f"  ⚠ Error al leer hoja 'clientes': {e}")

    try:
        df_gen = leer_excel(ruta_homologacion, p, sheet_name="generadores")
        df_gen = df_gen.dropna(subset=["NOMBRE_VARIANTE"])
        for _, row in df_gen.iterrows():
            variante = normalizar_nombre_para_buscar(row["NOMBRE_VARIANTE"])
            nombre_correcto = str(row["NOMBRE_CORRECTO"]).strip() if pd.notna(row["NOMBRE_CORRECTO"]) else None
            if variante and nombre_correcto:
                homolog_generadores[variante] = nombre_correcto
        p(f"  Generadores homologados: {len(homolog_generadores)} entradas")
    except Exception as e:
        p(f"  ⚠ Error al leer hoja 'generadores': {e}")

    return homolog_clientes, homolog_generadores


def aplicar_homologacion_cliente(df, col_nombre, col_rut, homolog_clientes):
    if not homolog_clientes:
        return df

    def corregir(row):
        variante = normalizar_nombre_para_buscar(row[col_nombre])
        if variante in homolog_clientes:
            correccion = homolog_clientes[variante]
            nuevo_nombre = correccion["nombre"]
            nuevo_rut = correccion["rut"] if correccion["rut"] else row[col_rut]
            return pd.Series([nuevo_nombre, nuevo_rut])
        return pd.Series([row[col_nombre], row[col_rut]])

    df[[col_nombre, col_rut]] = df.apply(corregir, axis=1)
    return df


def aplicar_homologacion_generador(df, col_generador, homolog_generadores):
    if not homolog_generadores:
        return df

    def corregir(valor):
        variante = normalizar_nombre_para_buscar(valor)
        return homolog_generadores.get(variante, valor)

    df[col_generador] = df[col_generador].apply(corregir)
    return df


def guardar_excel(df, ruta, p):
    for intento in range(1, MAX_REINTENTOS + 1):
        try:
            if ruta.exists():
                wb = load_workbook(ruta)
                if HOJA_DESTINO in wb.sheetnames:
                    del wb[HOJA_DESTINO]
                ws = wb.create_sheet(HOJA_DESTINO)
            else:
                wb = Workbook()
                ws = wb.active
                ws.title = HOJA_DESTINO

            ws.append(list(df.columns))
            for fila in df.itertuples(index=False):
                ws.append([None if str(v) in ("nan", "NaT", "None") else v for v in fila])

            wb.save(ruta)
            p(f"  ✓ Guardado en: {ruta.name}")
            return
        except PermissionError:
            p(f"  ⚠ Archivo bloqueado — intento {intento}/{MAX_REINTENTOS}")
            if intento < MAX_REINTENTOS:
                time.sleep(ESPERA_REINTENTO)

    raise PermissionError(f"No se pudo guardar {ruta.name}. Ciérralo en Excel.")


# ══════════════════════════════════════════════════════════════════
# LOOKUPS
# ══════════════════════════════════════════════════════════════════
def cargar_lookups(rutas, p):
    p("\n── Cargando tablas de lookup ──")

    df_t = leer_excel(rutas["transportistas"], p, sheet_name="Hoja1", header=0)
    df_t = df_t[["Transportista", "RUT"]].dropna(subset=["Transportista"])
    df_t["Transportista"] = df_t["Transportista"].astype(str).str.strip()
    df_t["RUT"] = df_t["RUT"].astype(str).str.strip()
    transportistas = dict(zip(df_t["Transportista"], df_t["RUT"]))
    p(f"  Transportistas:     {len(transportistas)} entradas")

    df_m = leer_excel(rutas["sinader"], p, sheet_name="Clasificación Residuos", header=1)
    df_m = df_m[["TIPO", "CÓDIGOS SINADER"]].dropna(subset=["TIPO"])
    df_m["TIPO"] = df_m["TIPO"].astype(str).str.strip()
    df_m["CÓDIGOS SINADER"] = df_m["CÓDIGOS SINADER"].astype(str).str.strip()
    df_m = df_m.drop_duplicates(subset=["TIPO"])
    materiales = dict(zip(df_m["TIPO"], df_m["CÓDIGOS SINADER"]))
    p(f"  Materiales SINADER: {len(materiales)} tipos")

    df_d = leer_excel(rutas["destinatarios"], p, sheet_name=HOJA_DESTINATARIOS, header=1)
    cols = [
        "NOMBRE DE FANTASÍA", "COMUNA DE ESTABLECIMIENTO", "CÓDIGOS LER",
        "CÓDIGO ESTABLECIMIENTO SINADER", "RUT DESTINATARIO",
        "CÓDIGO DE TRATAMIENTO SINADER",
    ]
    df_d = df_d[cols].dropna(subset=["NOMBRE DE FANTASÍA"])
    for c in cols:
        df_d[c] = df_d[c].astype(str).str.strip()
    p(f"  Destinatarios:      {len(df_d)} entradas")

    return transportistas, materiales, df_d


# ══════════════════════════════════════════════════════════════════
# FUENTES
# ══════════════════════════════════════════════════════════════════
def procesar_bo(rutas, homolog_c, homolog_g, p):
    p("\n── BO San Bernardo ──")
    df = leer_tabla_nombrada(rutas["bo"], "Tabla_operacion", p)
    n_leidas = len(df)
    p(f"  • Filas leídas: {n_leidas}")

    # Filtro año
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    n_antes = len(df)
    df = df[df["Fecha"].dt.year >= AÑO_DESDE]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro año ≥ {AÑO_DESDE}: {len(df)} filas (excluidas: {n_excluidas})")

    df["Movimiento"] = "Traslado"
    df["Gestor"] = "AMBIPAR ENVIRONMENT CHILE"
    df["Transportista"] = "AMBIPAR ENVIRONMENT CHILE"
    df["Rut transportista"] = "96824110-9"
    df["Destino inferido"] = "No"
    df["Región"] = REGION_RM
    df["Movimiento interempresa"] = "No"  # el concepto solo aplica entre Ecofibras SB y GIRI
    df = df.rename(columns={"Comuna destino": "Comuna Destino"})

    # Filtro destino
    n_antes = len(df)
    df = df[df["Destino"].notna()]
    df = df[~df["Destino"].isin(["ECOFIBRAS SAN BERNARDO", "GIRI"])]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro Destino válido: {len(df)} filas (excluidas: {n_excluidas})")

    # Homologación
    n_cli_antes = df["Cliente"].nunique()
    df = aplicar_homologacion_cliente(df, "Cliente", "RUT", homolog_c)
    n_cli_despues = df["Cliente"].nunique()
    p(f"  • Homologación aplicada: clientes únicos {n_cli_antes} → {n_cli_despues}")

    df = aplicar_homologacion_generador(df, "Generador", homolog_g)

    # Alertas
    sin_patente = df[df["Patente de Camión"].isna()]
    if len(sin_patente) > 0:
        p(f"  ⚠ Filas sin patente: {len(sin_patente)} (SE MANTIENEN en el consolidado)")

    df = df[[
        "Fecha", "Mes", "Cliente", "RUT", "Gestor", "Contrato", "Generador",
        "Transportista", "Rut transportista", "Patente de Camión", "Ticket de pesaje",
        "Peso neto", "Unidad", "Destino", "Comuna Destino", "TIPO", "Movimiento",
        "Región", "Destino inferido", "Movimiento interempresa",
    ]]
    p(f"  ✓ BO final: {len(df)} filas")
    return df


def procesar_bo_valparaiso(rutas, homolog_c, homolog_g, p):
    p("\n── BO Valparaíso ──")
    df = leer_tabla_nombrada(rutas["bo_valparaiso"], "Tabla_operacion", p)
    n_leidas = len(df)
    p(f"  • Filas leídas: {n_leidas}")

    # Ajustes de columnas propios de esta fuente (no calzan 1 a 1 con BO San Bernardo)
    df = df.rename(columns={
        "Peso neto [kg]": "Peso neto",
        "Comuna destino": "Comuna Destino",
    })
    df["Unidad"] = "kg"

    # Filtro año
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    n_antes = len(df)
    df = df[df["Fecha"].dt.year >= AÑO_DESDE]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro año ≥ {AÑO_DESDE}: {len(df)} filas (excluidas: {n_excluidas})")

    df["Movimiento"] = "Traslado"
    df["Gestor"] = "AMBIPAR ENVIRONMENT CHILE"
    df["Transportista"] = "AMBIPAR ENVIRONMENT CHILE"
    df["Rut transportista"] = "96824110-9"
    df["Destino inferido"] = "No"
    df["Región"] = REGION_VALPARAISO
    df["Movimiento interempresa"] = "No"  # BO Valparaíso no traslada a plantas propias

    # Filtro destino (sin exclusiones internas: no aplica a Ecofibras/GIRI)
    n_antes = len(df)
    df = df[df["Destino"].notna()]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro Destino no vacío: {len(df)} filas (excluidas: {n_excluidas})")

    # Homologación
    n_cli_antes = df["Cliente"].nunique()
    df = aplicar_homologacion_cliente(df, "Cliente", "RUT", homolog_c)
    n_cli_despues = df["Cliente"].nunique()
    p(f"  • Homologación aplicada: clientes únicos {n_cli_antes} → {n_cli_despues}")

    df = aplicar_homologacion_generador(df, "Generador", homolog_g)

    # Alertas
    sin_patente = df[df["Patente de Camión"].isna()]
    if len(sin_patente) > 0:
        p(f"  ⚠ Filas sin patente: {len(sin_patente)} (SE MANTIENEN en el consolidado)")

    df = df[[
        "Fecha", "Mes", "Cliente", "RUT", "Gestor", "Contrato", "Generador",
        "Transportista", "Rut transportista", "Patente de Camión", "Ticket de pesaje",
        "Peso neto", "Unidad", "Destino", "Comuna Destino", "TIPO", "Movimiento",
        "Región", "Destino inferido", "Movimiento interempresa",
    ]]
    p(f"  ✓ BO Valparaíso final: {len(df)} filas")
    return df


def procesar_giri(rutas, transportistas, homolog_c, homolog_g, p):
    p("\n── GIRI ──")
    df = leer_tabla_nombrada(rutas["giri"], "INGRESO", p)
    n_leidas = len(df)
    p(f"  • Filas leídas: {n_leidas}")

    # Normalizar y filtrar Movimiento (la tabla trae filas vacías/fila Total al final)
    df["MOVIMIENTO"] = df["MOVIMIENTO"].astype(str).str.strip().str.upper()
    n_antes = len(df)
    df = df[df["MOVIMIENTO"].isin(["INGRESO", "TRASLADO"])]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro Movimiento válido (Ingreso/Traslado): {len(df)} filas (excluidas: {n_excluidas})")

    # Filtro año
    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    n_antes = len(df)
    df = df[df["FECHA"].dt.year >= AÑO_DESDE]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro año ≥ {AÑO_DESDE}: {len(df)} filas (excluidas: {n_excluidas})")

    # Filtro cliente
    n_antes = len(df)
    df = df[df["CLIENTE"].notna()]
    df = df[~df["CLIENTE"].astype(str).str.strip().isin(["0"])]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro Cliente válido: {len(df)} filas (excluidas: {n_excluidas})")

    df["Unidad"] = "kg"

    es_traslado = df["MOVIMIENTO"].eq("TRASLADO")

    # Destino: si es Traslado, toma la columna DESTINO de la planilla;
    # si es Ingreso, sigue siendo GIRI por defecto (como antes)
    df["Destino"] = df["DESTINO"].astype(str).str.strip().where(es_traslado, "GIRI")

    # Comuna Destino: fija para Ingreso (QUILICURA); para Traslado se deja vacía
    # a propósito — el join global con BBDD DESTINATARIO.xlsx (más abajo en el
    # pipeline) ya la completa automáticamente cruzando por nombre de Destino.
    df["Comuna Destino"] = pd.NA
    df.loc[~es_traslado, "Comuna Destino"] = "QUILICURA"

    df["Movimiento"] = es_traslado.map({True: "Traslado", False: "Ingreso"})

    # Excluir traslados con Destino=ECOFIBRAS SAN BERNARDO: ese movimiento ya
    # queda registrado con su propio ticket de pesaje en el libro de Ingreso de
    # Ecofibras (CLIENTE=GIRI). Dejar ambas filas duplicaría el peso — mismo
    # criterio que usa BO San Bernardo y que se aplica en procesar_ecofibras().
    n_antes = len(df)
    df = df[~((es_traslado) & (df["Destino"].str.upper() == "ECOFIBRAS SAN BERNARDO"))]
    es_traslado = es_traslado.loc[df.index]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro Traslado→Ecofibras excluido (ya registrado en Ecofibras): {len(df)} filas (excluidas: {n_excluidas})")

    n_traslados = int(es_traslado.sum())
    p(f"  • Movimientos: {len(df) - n_traslados} Ingreso / {n_traslados} Traslado")

    df["Destino inferido"] = "No"
    df["Región"] = REGION_RM

    # Interempresa: Ecofibras → GIRI (ingreso con Cliente=ECOFIBRAS S.A).
    # El caso GIRI → Ecofibras ya no aparece aquí, se excluyó arriba para no duplicar.
    es_interempresa = df["CLIENTE"].astype(str).str.strip().eq("ECOFIBRAS S.A") & ~es_traslado
    df["Movimiento interempresa"] = es_interempresa.map({True: "Sí", False: "No"})
    n_interempresa = int(es_interempresa.sum())
    p(f"  • Movimientos interempresa (Ecofibras → GIRI): {n_interempresa}")

    df["Transportista"] = df["GESTOR"].astype(str).str.strip()
    df["Rut transportista"] = df["Transportista"].map(transportistas)

    # Alertas transportista
    sin_rut = df[df["Rut transportista"].isna()]["Transportista"].dropna().unique()
    if len(sin_rut) > 0:
        p(f"  ⚠ Gestores sin RUT en Transportistas: {len(sin_rut)}")
        for g in sin_rut[:5]:
            p(f"     → '{g}'")

    df = df.rename(columns={
        "FECHA": "Fecha",
        "MES": "Mes",
        "CLIENTE": "Cliente",
        "RUT CLIENTE": "RUT",
        "GESTOR": "Gestor",
        "GENERADOR": "Generador",
        "N° CONTRATO": "Contrato",
        "PATENTE DE CAMIÓN": "Patente de Camión",
        "TICKET DE PESAJE": "Ticket de pesaje",
        "PESO NETO KG": "Peso neto",
    })

    # Homologación
    n_cli_antes = df["Cliente"].nunique()
    df = aplicar_homologacion_cliente(df, "Cliente", "RUT", homolog_c)
    n_cli_despues = df["Cliente"].nunique()
    p(f"  • Homologación aplicada: clientes únicos {n_cli_antes} → {n_cli_despues}")

    df = aplicar_homologacion_generador(df, "Generador", homolog_g)

    df = df[[
        "Fecha", "Mes", "Cliente", "RUT", "Gestor", "Contrato", "Generador",
        "Transportista", "Rut transportista", "Patente de Camión", "Ticket de pesaje",
        "Peso neto", "Unidad", "Destino", "Comuna Destino", "TIPO", "Movimiento",
        "Región", "Destino inferido", "Movimiento interempresa",
    ]]
    p(f"  ✓ GIRI final: {len(df)} filas")
    return df


def procesar_ecofibras(rutas, transportistas, homolog_c, homolog_g, p):
    p("\n── Ecofibras San Bernardo ──")
    df = leer_tabla_nombrada(rutas["ecofibras"], "INGRESO", p)
    n_leidas = len(df)
    p(f"  • Filas leídas: {n_leidas}")

    # Normalizar y filtrar Movimiento (la tabla trae filas vacías al final)
    df["MOVIMIENTO"] = df["MOVIMIENTO"].astype(str).str.strip().str.upper()
    n_antes = len(df)
    df = df[df["MOVIMIENTO"].isin(["INGRESO", "TRASLADO"])]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro Movimiento válido (Ingreso/Traslado): {len(df)} filas (excluidas: {n_excluidas})")

    # Filtro año
    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    n_antes = len(df)
    df = df[df["FECHA"].dt.year >= AÑO_DESDE]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro año ≥ {AÑO_DESDE}: {len(df)} filas (excluidas: {n_excluidas})")

    # Filtro generador
    n_antes = len(df)
    df = df[df["GENERADOR"].astype(str).str.strip() != "REMBRE SPA"]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro Generador ≠ 'REMBRE SPA': {len(df)} filas (excluidas: {n_excluidas})")

    df["Unidad"] = "kg"

    es_traslado = df["MOVIMIENTO"].eq("TRASLADO")

    # Destino: si es Traslado, toma la columna DESTINO de la planilla;
    # si es Ingreso, sigue siendo ECOFIBRAS SAN BERNARDO por defecto (como antes)
    df["Destino"] = df["DESTINO"].astype(str).str.strip().where(es_traslado, "ECOFIBRAS SAN BERNARDO")

    # Comuna Destino: fija para Ingreso (SAN BERNARDO); para Traslado se deja vacía
    # a propósito — el join global con BBDD DESTINATARIO.xlsx más adelante en el
    # pipeline ya la completa automáticamente cruzando por nombre de Destino.
    df["Comuna Destino"] = pd.NA
    df.loc[~es_traslado, "Comuna Destino"] = "SAN BERNARDO"

    df["Movimiento"] = es_traslado.map({True: "Traslado", False: "Ingreso"})

    # Excluir traslados con Destino=GIRI: ese movimiento ya queda registrado
    # con su propio ticket de pesaje en el libro de Ingreso de GIRI (CLIENTE=ECOFIBRAS S.A).
    # Dejar ambas filas duplicaría el peso — mismo criterio que usa BO San Bernardo.
    n_antes = len(df)
    df = df[~((es_traslado) & (df["Destino"].str.upper() == "GIRI"))]
    es_traslado = es_traslado.loc[df.index]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro Traslado→GIRI excluido (ya registrado en GIRI): {len(df)} filas (excluidas: {n_excluidas})")

    n_traslados = int(es_traslado.sum())
    p(f"  • Movimientos: {len(df) - n_traslados} Ingreso / {n_traslados} Traslado")

    df["Destino inferido"] = "No"
    df["Región"] = REGION_RM
    df["Transportista"] = df["GESTOR"].astype(str).str.strip()
    df["Rut transportista"] = df["Transportista"].map(transportistas)

    # Interempresa: GIRI → Ecofibras (ingreso con Cliente=GIRI).
    # El caso Ecofibras → GIRI ya no aparece aquí, se excluyó arriba para no duplicar.
    es_interempresa = df["CLIENTE"].astype(str).str.strip().str.upper().eq("GIRI") & ~es_traslado
    df["Movimiento interempresa"] = es_interempresa.map({True: "Sí", False: "No"})
    n_interempresa = int(es_interempresa.sum())
    p(f"  • Movimientos interempresa (GIRI → Ecofibras): {n_interempresa}")

    # Alertas transportista
    sin_rut = df[df["Rut transportista"].isna()]["Transportista"].dropna().unique()
    if len(sin_rut) > 0:
        p(f"  ⚠ Gestores sin RUT en Transportistas: {len(sin_rut)}")
        for g in sin_rut[:5]:
            p(f"     → '{g}'")

    df = df.rename(columns={
        "FECHA": "Fecha",
        "MES": "Mes",
        "CLIENTE": "Cliente",
        "RUT CLIENTE": "RUT",
        "GESTOR": "Gestor",
        "GENERADOR": "Generador",
        "N° CONTRATO": "Contrato",
        "PATENTE CAMIÓN": "Patente de Camión",
        "TICKET DE PESAJE": "Ticket de pesaje",
        "PESO NETO KG": "Peso neto",
    })

    # Homologación
    n_cli_antes = df["Cliente"].nunique()
    df = aplicar_homologacion_cliente(df, "Cliente", "RUT", homolog_c)
    n_cli_despues = df["Cliente"].nunique()
    p(f"  • Homologación aplicada: clientes únicos {n_cli_antes} → {n_cli_despues}")

    df = aplicar_homologacion_generador(df, "Generador", homolog_g)

    df = df[[
        "Fecha", "Mes", "Cliente", "RUT", "Gestor", "Contrato", "Generador",
        "Transportista", "Rut transportista", "Patente de Camión", "Ticket de pesaje",
        "Peso neto", "Unidad", "Destino", "Comuna Destino", "TIPO", "Movimiento",
        "Región", "Destino inferido", "Movimiento interempresa",
    ]]
    p(f"  ✓ Ecofibras final: {len(df)} filas")
    return df


def procesar_proveedores(rutas, transportistas, homolog_c, homolog_g, p):
    p("\n── Proveedores ──")
    df = leer_tabla_nombrada(rutas["proveedores"], "Tabla2", p)
    n_leidas = len(df)
    p(f"  • Filas leídas: {n_leidas}")

    # Filtro año
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    n_antes = len(df)
    df = df[df["Fecha"].dt.year >= AÑO_DESDE]
    n_excluidas = n_antes - len(df)
    p(f"  • Filtro año ≥ {AÑO_DESDE}: {len(df)} filas (excluidas: {n_excluidas})")

    df = df.rename(columns={
        "RUT CLIENTE": "RUT",
        "N° Guía": "Ticket de pesaje",
        "Cantidad (Kg)": "Peso neto",
        "Patente": "Patente de Camión",
        "Destinatario": "Destino",
    })
    df["Unidad"] = "kg"
    df["Movimiento"] = "Traslado externo"
    df["Contrato"] = None
    df["Destino inferido"] = "No"
    df["Región"] = REGION_RM
    df["Movimiento interempresa"] = "No"
    df["Transportista"] = df["Gestor"].astype(str).str.strip()
    df["Rut transportista"] = df["Transportista"].map(transportistas)

    # Alertas transportista
    sin_rut = df[df["Rut transportista"].isna()]["Transportista"].dropna().unique()
    if len(sin_rut) > 0:
        p(f"  ⚠ Gestores sin RUT en Transportistas: {len(sin_rut)}")
        for g in sin_rut[:5]:
            p(f"     → '{g}'")

    # Homologación
    n_cli_antes = df["Cliente"].nunique()
    df = aplicar_homologacion_cliente(df, "Cliente", "RUT", homolog_c)
    n_cli_despues = df["Cliente"].nunique()
    p(f"  • Homologación aplicada: clientes únicos {n_cli_antes} → {n_cli_despues}")

    df = aplicar_homologacion_generador(df, "Generador", homolog_g)

    df = df[[
        "Fecha", "Mes", "Cliente", "RUT", "Gestor", "Contrato", "Generador",
        "Transportista", "Rut transportista", "Patente de Camión", "Ticket de pesaje",
        "Peso neto", "Unidad", "Destino", "Comuna Destino", "TIPO", "Movimiento",
        "Región", "Destino inferido", "Movimiento interempresa",
    ]]
    p(f"  ✓ Proveedores final: {len(df)} filas")
    return df


def procesar_tradicionales(rutas, homolog_c, p):
    p("\n── Tradicionales ──")
    p("  • Fuente: tabla 'Analizado' ya actualizada en BBDD TRADICIONALES.xlsx")

    # ⚠️ REGISTRO INICIAL — GARANTÍA DE NO PERDER FILAS
    df = leer_tabla_nombrada(rutas["tradicionales"], "Analizado", p)
    n_leidas = len(df)
    p(f"  • Filas leídas de 'Analizado': {n_leidas}")

    # Tipos base
    df["Fec_salida"] = pd.to_datetime(df["Fec_salida"], errors="coerce")
    df["Rut"] = pd.to_numeric(df["Rut"], errors="coerce")
    df["Contratato"] = pd.to_numeric(df["Contratato"], errors="coerce")
    df["Peso Final [kg]"] = pd.to_numeric(df["Peso Final [kg]"], errors="coerce")
    df["RtHruFol"] = pd.to_numeric(df["RtHruFol"], errors="coerce")

    if "Camiones.Patente" in df.columns:
        df["Camiones.Patente"] = limpiar_texto_serie(df["Camiones.Patente"])
    else:
        df["Camiones.Patente"] = None

    if "Destino" in df.columns:
        df["Destino"] = limpiar_texto_serie(df["Destino"])
    else:
        df["Destino"] = None

    # Filtro año
    n_antes = len(df)
    df = df[df["Fec_salida"].dt.year >= AÑO_DESDE]
    p(f"  • Filtro año ≥ {AÑO_DESDE}: {len(df)} filas  (excluidas: {n_antes - len(df)})")

    # Renombres / columnas derivadas
    df = df.rename(columns={
        "Fec_salida": "Fecha",
        "Raz_social": "Cliente",
        "Rut": "RUT",
        "Contratato": "Contrato",
        "RtHruFol": "Ticket de pesaje",
        "Peso Final [kg]": "Peso neto",
        "Camiones.Patente": "Patente de Camión",
    })

    df["Mes"] = df["Fecha"].apply(nombre_mes)
    df["Gestor"] = "AMBIPAR ENVIRONMENT CHILE"
    df["Generador"] = df["Cliente"]
    df["Transportista"] = "AMBIPAR ENVIRONMENT CHILE"
    df["Rut transportista"] = "96824110-9"
    df["Unidad"] = "kg"
    df["TIPO"] = "ASIMILABLE A DOMICILIARIO"
    df["Movimiento"] = "Traslado"
    df["Destino inferido"] = "No"
    df["Región"] = REGION_RM
    df["Movimiento interempresa"] = "No"

    # Reemplazos equivalentes a Power Query
    reemplazos_cliente = {
        "GESTION VIAL S.A.": "GESTION VIAL S.A",
        "INMOBILIARIA RECONQUISTA S.A.": "INMOBILIARIA RECONQUISTA S.A",
        "INMOBILIARIA Y CONSTRUCTORA PEDRO DE VALDIVIA II S.A.": "INMOBILIARIA Y CONSTRUCTORA PEDRO DE VALDIVIA II S.A",
        "KOMATSU CHILE S.A.": "KOMATSU CHILE S.A",
        "LINDE GAS CHILE S.A.": "LINDE GAS CHILE S.A",
        "SACYR CHILE SA": "SACYR CHILE S.A",
        "SERVICIOS AVO II SpA": "SERVICIOS AVO II SPA",
        "SODEXO CHILE SPA.": "SODEXO CHILE SPA",
        "Servicios COVI SpA": "SERVICIOS COVI SPA",
    }
    df["Cliente"] = df["Cliente"].replace(reemplazos_cliente)

    # Filtros como en la consulta final de trazabilidad
    n_antes = len(df)
    df = df[df["estado"] == "C"]
    p(f"  • Filtro estado = 'C': {len(df)} filas  (excluidas: {n_antes - len(df)})")

    n_antes = len(df)
    df = df[df["Peso neto"].fillna(0) != 0]
    p(f"  • Filtro Peso neto ≠ 0: {len(df)} filas  (excluidas: {n_antes - len(df)})")

    # Homologación ANTES de inferir destino
    n_cli_antes = df["Cliente"].nunique(dropna=True)
    df = aplicar_homologacion_cliente(df, "Cliente", "RUT", homolog_c)
    n_cli_despues = df["Cliente"].nunique(dropna=True)
    p(f"  • Homologación aplicada: clientes únicos {n_cli_antes} → {n_cli_despues}")

    # ⚠️ ALERTA: Filas incompletas (sin patente, sin destino, sin ticket)
    sin_patente = df[df["Patente de Camión"].isna()]
    sin_ticket = df[df["Ticket de pesaje"].isna()]
    sin_destino = df[df["Destino"].isna()]
    if len(sin_patente) > 0:
        p(f"  ⚠ Filas sin patente: {len(sin_patente)} (SE MANTIENEN en el consolidado)")
    if len(sin_ticket) > 0:
        p(f"  ⚠ Filas sin ticket de pesaje: {len(sin_ticket)} (SE MANTIENEN en el consolidado)")
    if len(sin_destino) > 0:
        p(f"  ⚠ Filas sin destino: {len(sin_destino)} (se van a inferir o marcar pendiente)")

    # Inferencia de destino SOLO si está vacío
    df["Destino"] = limpiar_texto_serie(df["Destino"])

    df_con_dest = df[df["Destino"].notna()].copy()
    destino_frecuente = (
        df_con_dest.groupby("Cliente")["Destino"]
        .agg(lambda x: x.value_counts().index[0] if len(x) > 0 else None)
        .to_dict()
    )

    def resolver_destino(row):
        destino_actual = row["Destino"]
        if pd.notna(destino_actual) and str(destino_actual).strip() not in ("", "None", "nan"):
            return pd.Series([destino_actual, "No"])

        destino_hist = destino_frecuente.get(row["Cliente"])
        if pd.notna(destino_hist) and str(destino_hist).strip() not in ("", "None", "nan"):
            return pd.Series([destino_hist, "Sí (frecuencia)"])

        # Sin fallback automático: si no hay destino confirmado, marcar pendiente
        return pd.Series([None, "Pendiente revisión"])

    res = df.apply(resolver_destino, axis=1)
    df["Destino"] = res[0]
    df["Destino inferido"] = res[1]

    n_inferidos = (df["Destino inferido"] == "Sí (frecuencia)").sum()
    n_pendientes = (df["Destino inferido"] == "Pendiente revisión").sum()
    p(f"  • Destinos inferidos por frecuencia: {n_inferidos}")
    p(f"  • Destinos pendientes de revisión:   {n_pendientes}")

    # La comuna se completa después con Destinatarios
    df["Comuna Destino"] = None

    # ⚠️ GARANTÍA: Seleccionar columnas SIN perder filas
    df = df[[
        "Fecha", "Mes", "Cliente", "RUT", "Gestor", "Contrato", "Generador",
        "Transportista", "Rut transportista", "Patente de Camión", "Ticket de pesaje",
        "Peso neto", "Unidad", "Destino", "Comuna Destino", "TIPO", "Movimiento",
        "Región", "Destino inferido", "Movimiento interempresa",
    ]]

    # ⚠️ VALIDACIÓN CRÍTICA: ¿Se perdieron filas?
    n_finales = len(df)
    if n_finales < n_leidas:
        p(f"  ⚠ ALERTA: Se perdieron {n_leidas - n_finales} filas")
        p(f"     Entrada: {n_leidas} | Salida: {n_finales}")
        p(f"     (Esto no debería ocurrir — revisar filtros)")
    elif n_finales == n_leidas:
        p(f"  ✓ Garantía cumplida: 0 filas perdidas ({n_finales} = {n_leidas})")
    else:
        p(f"  ⚠ Filas ganadas (?): {n_finales} > {n_leidas}")

    p(f"  ✓ Tradicionales final: {n_finales} filas")
    return df


# ══════════════════════════════════════════════════════════════════
# PIPELINE
# ══════════════════════════════════════════════════════════════════
def consolidar(rutas, ruta_prueba=None, ruta_log=None, modo_reset=False,
               modo_prueba=False, mostrar=True):
    """Ejecuta la consolidación completa de RM.

    Parámetros
    ----------
    rutas : dict
        Diccionario con las 9 rutas de fuentes/tablas + 'homologacion' y
        'destino_real'. Se arma con rutas_desde_config().
    ruta_prueba : str o Path, opcional
        Dónde escribir el consolidado en modo prueba. Obligatorio si
        modo_prueba=True.
    ruta_log : str o Path, opcional
        Archivo de auditoría (se agrega, nunca se borra). Si se omite, no
        se escribe ningún log en disco — solo queda en el 'log' devuelto.
    modo_reset : bool
        Si es True, reemplaza la trazabilidad completa en vez de acumular.
    modo_prueba : bool
        Si es True, escribe en ruta_prueba y NO toca el archivo real.
    mostrar : bool
        Si es True imprime el avance en pantalla.

    Devuelve
    --------
    dict con las claves:
        consolidado   : DataFrame final (mismo que se escribió)
        filas, columnas
        log           : todo el texto del avance
    """
    p = Registro(mostrar=mostrar, ruta_log=ruta_log)

    verificar_rutas(rutas)
    if modo_prueba and not ruta_prueba:
        raise ValueError("modo_prueba=True requiere indicar ruta_prueba.")

    p("\n" + "═" * 60)
    if modo_prueba:
        p("  INICIO — MODO PRUEBA")
    elif modo_reset:
        p("  INICIO — RESET COMPLETO")
    else:
        p("  INICIO — MODO ACUMULATIVO")
    p(f"  {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    p(f"  Procesando registros desde: {AÑO_DESDE} en adelante")
    p("═" * 60)

    homolog_c, homolog_g = cargar_homologacion(rutas["homologacion"], p)
    transportistas, materiales, df_dest = cargar_lookups(rutas, p)

    df_bo = procesar_bo(rutas, homolog_c, homolog_g, p)
    df_bo_valp = procesar_bo_valparaiso(rutas, homolog_c, homolog_g, p)
    df_giri = procesar_giri(rutas, transportistas, homolog_c, homolog_g, p)
    df_eco = procesar_ecofibras(rutas, transportistas, homolog_c, homolog_g, p)
    df_prov = procesar_proveedores(rutas, transportistas, homolog_c, homolog_g, p)
    df_trad = procesar_tradicionales(rutas, homolog_c, p)

    p("\n── Combinando las 6 fuentes ──")
    conteos = {
        "BO": len(df_bo),
        "BO_Valparaíso": len(df_bo_valp),
        "GIRI": len(df_giri),
        "Ecofibras": len(df_eco),
        "Proveedores": len(df_prov),
        "Tradicionales": len(df_trad),
    }
    for nombre, n in conteos.items():
        p(f"  {nombre:<15} {n:>6} filas")
    suma_fuentes = sum(conteos.values())
    p(f"  {'─' * 24}")
    p(f"  Suma esperada:   {suma_fuentes:>6} filas")

    df_bo["__origen__"] = "BO"
    df_bo_valp["__origen__"] = "BO_Valparaíso"
    df_giri["__origen__"] = "GIRI"
    df_eco["__origen__"] = "Ecofibras"
    df_prov["__origen__"] = "Proveedores"
    df_trad["__origen__"] = "Tradicionales"

    df = pd.concat([df_bo, df_bo_valp, df_giri, df_eco, df_prov, df_trad], ignore_index=True)
    df["__idx__"] = df.index

    # ⚠️ VALIDACIÓN CRÍTICA: ¿Se perdieron filas en la combinación?
    suma_combinada = len(df)
    if suma_combinada == suma_fuentes:
        p(f"  ✓ Combinado: {suma_combinada} filas (suma exacta, 0 pérdidas)")
    elif suma_combinada < suma_fuentes:
        p(f"  ⚠ ALERTA: Se perdieron filas en la combinación")
        p(f"     Esperadas: {suma_fuentes} | Obtenidas: {suma_combinada}")
        p(f"     Diferencia: {suma_fuentes - suma_combinada} filas PERDIDAS")
    else:
        p(f"  ⚠ Filas ganadas (?): {suma_combinada} > {suma_fuentes}")

    # ── Join SINADER ─────────────────────────────────────────────
    p("\n── Join con SINADER (TIPO → CÓDIGOS SINADER) ──")
    df["CÓDIGOS SINADER"] = df["TIPO"].astype(str).str.strip().map(materiales)

    sin_tipo = df[df["CÓDIGOS SINADER"].isna() & df["TIPO"].notna()][["__origen__", "TIPO"]]
    if len(sin_tipo) > 0:
        agrupado_tipo = (
            sin_tipo.groupby("TIPO")["__origen__"]
            .agg(lambda x: ", ".join(sorted(set(x))))
            .reset_index()
            .rename(columns={"__origen__": "ORIGEN"})
        )
        p(f"  ⚠ {len(sin_tipo)} filas sin código SINADER:")
        for _, r in agrupado_tipo.head(15).iterrows():
            p(f"     [{r['ORIGEN']}] → '{r['TIPO']}'")
    else:
        p("  ✓ Todos los TIPO tienen código SINADER")

    # ── Join DESTINATARIOS ──────────────────────────────────────
    p("\n── Join con DESTINATARIOS (Destino+Comuna+Código LER) ──")

    lookup_comuna = (
        df_dest[["NOMBRE DE FANTASÍA", "COMUNA DE ESTABLECIMIENTO"]]
        .drop_duplicates(subset=["NOMBRE DE FANTASÍA"])
    )
    mask_sin = df["Comuna Destino"].isna()
    if mask_sin.any():
        df_sin = df[mask_sin][["Destino"]].merge(
            lookup_comuna,
            left_on="Destino",
            right_on="NOMBRE DE FANTASÍA",
            how="left",
        )
        df.loc[mask_sin, "Comuna Destino"] = df_sin["COMUNA DE ESTABLECIMIENTO"].values

    df["_k_dest"] = df["Destino"].astype(str).str.strip()
    df["_k_comuna"] = df["Comuna Destino"].astype(str).str.strip()
    df["_k_ler"] = df["CÓDIGOS SINADER"].astype(str).str.strip()

    df_dj = df_dest.copy()
    for c in ["NOMBRE DE FANTASÍA", "COMUNA DE ESTABLECIMIENTO", "CÓDIGOS LER"]:
        df_dj[c] = df_dj[c].astype(str).str.strip()

    n_antes = len(df)
    df = df.merge(
        df_dj[[
            "NOMBRE DE FANTASÍA", "COMUNA DE ESTABLECIMIENTO", "CÓDIGOS LER",
            "CÓDIGO ESTABLECIMIENTO SINADER", "RUT DESTINATARIO",
            "CÓDIGO DE TRATAMIENTO SINADER",
        ]],
        left_on=["_k_dest", "_k_comuna", "_k_ler"],
        right_on=["NOMBRE DE FANTASÍA", "COMUNA DE ESTABLECIMIENTO", "CÓDIGOS LER"],
        how="left",
    )
    df = df.drop_duplicates(subset=["__idx__"])
    df = df.drop(columns=[
        "__idx__", "_k_dest", "_k_comuna", "_k_ler",
        "NOMBRE DE FANTASÍA", "COMUNA DE ESTABLECIMIENTO", "CÓDIGOS LER",
    ], errors="ignore")

    simbolo = "✓" if len(df) == n_antes else "⚠"
    p(f"  {simbolo} JOIN: {n_antes} → {len(df)} filas")

    sin_dest_full = df[df["RUT DESTINATARIO"].isna()][
        ["__origen__", "Destino", "Comuna Destino", "CÓDIGOS SINADER"]
    ].copy()
    if len(sin_dest_full) > 0:
        agrupado_dest = (
            sin_dest_full.groupby(["Destino", "Comuna Destino", "CÓDIGOS SINADER"])["__origen__"]
            .agg(lambda x: ", ".join(sorted(set(x))))
            .reset_index()
            .rename(columns={"__origen__": "ORIGEN"})
        )
        p(f"  ⚠ {len(agrupado_dest)} combinaciones sin match en Destinatarios:")
        for _, r in agrupado_dest.head(15).iterrows():
            p(f"     [{r['ORIGEN']}] '{r['Destino']}' | '{r['Comuna Destino']}' | '{r['CÓDIGOS SINADER']}'")
    else:
        p("  ✓ Todos los destinos encontrados en tabla Destinatarios")

    # Limpieza final
    df = df.drop(columns=["Unidad", "__origen__"], errors="ignore")
    df = df.rename(columns={"Peso neto": "Peso neto (kg)"})

    # Estandarizar mes en español.
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df["Mes"] = df["Fecha"].apply(nombre_mes)

    # Región: cada procesador ya asigna la suya (BO Valparaíso → Valparaíso).
    # Aquí SOLO se rellenan las que quedaron vacías, con RM por defecto.
    if "Región" not in df.columns:
        df["Región"] = None
    n_sin_region = int(df["Región"].isna().sum())
    df["Región"] = df["Región"].fillna(REGION_RM)
    if n_sin_region:
        p(f"  • Filas sin región asignada → {REGION_RM}: {n_sin_region}")
    for reg, n in df["Región"].value_counts().items():
        p(f"     {reg}: {n} filas")

    # Evitar que entren filas tipo total/arrastre sin datos mínimos
    df = limpiar_filas_invalidas(df, p)

    for col in COLUMNAS_FINALES:
        if col not in df.columns:
            df[col] = None
    df = df[COLUMNAS_FINALES]

    p(f"\n── DataFrame final: {len(df)} filas × {len(df.columns)} columnas ──")

    # Guardar
    if modo_prueba:
        ruta_prueba = Path(ruta_prueba)
        # Se agrega fecha y hora al nombre para que cada corrida quede aparte
        # y no pise la anterior. Esto es temporal para la Etapa 1: en la
        # Etapa 2 el destino del archivo lo decide la aplicación, no una
        # carpeta fija en este computador.
        sello = datetime.now().strftime("%Y-%m-%d_%H%M")
        ruta_prueba = ruta_prueba.with_name(f"{ruta_prueba.stem}_{sello}{ruta_prueba.suffix}")
        p(f"\n── Modo PRUEBA — guardando en: {ruta_prueba.name} ──")
        p("  (El archivo real de trazabilidad NO se modifica)")
        guardar_excel(df, ruta_prueba, p)
        df_final = df
    else:
        ruta_destino = Path(rutas["destino_real"])
        if modo_reset:
            p("\n── Modo RESET — reemplazando trazabilidad completa ──")
            df_final = df.copy()
        else:
            p("\n── Modo ACUMULATIVO ──")
            if ruta_destino.exists():
                df_existente = leer_excel(ruta_destino, p, sheet_name=HOJA_DESTINO)
                p(f"  Trazabilidad existente: {len(df_existente)} filas")
            else:
                df_existente = pd.DataFrame(columns=COLUMNAS_FINALES)
                p("  Archivo destino no existe — se creará nuevo")

            if not df_existente.empty:
                for col in CLAVE_DEDUP:
                    if col in df_existente.columns:
                        df_existente[col] = df_existente[col].astype(str)
                    if col in df.columns:
                        df[col] = df[col].astype(str)

                llave_exist = df_existente[CLAVE_DEDUP].apply(tuple, axis=1)
                llave_nueva = df[CLAVE_DEDUP].apply(tuple, axis=1)
                df_nuevas = df[~llave_nueva.isin(llave_exist)].copy()
            else:
                df_nuevas = df.copy()

            p(f"  Filas nuevas a agregar: {len(df_nuevas)}")
            p(f"  Duplicadas omitidas:    {len(df) - len(df_nuevas)}")
            df_final = pd.concat([df_existente, df_nuevas], ignore_index=True)

        p(f"  Total final: {len(df_final)} filas")
        guardar_excel(df_final, ruta_destino, p)

    p("\n" + "═" * 60)
    p("  CONSOLIDACIÓN COMPLETADA ✓")
    p("═" * 60 + "\n")

    return {
        "consolidado": df_final,
        "filas": len(df_final),
        "columnas": len(df_final.columns),
        "log": p.texto(),
    }


# ══════════════════════════════════════════════════════════════════
# USO DESDE LA TERMINAL — se conserva para poder seguir trabajando igual
# ══════════════════════════════════════════════════════════════════
def _main():
    """Permite ejecutar este archivo desde la terminal, igual que antes.

        python consolidar.py --prueba
        python consolidar.py --reset
        python consolidar.py
    """
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from config import carpetas, ConfiguracionFaltante

    modo_reset = "--reset" in sys.argv
    modo_prueba = "--prueba" in sys.argv

    try:
        rutas = rutas_desde_config()
        c = carpetas()
    except ConfiguracionFaltante as e:
        print(f"\n✗ {e}\n")
        sys.exit(2)

    ruta_prueba = c["rm"] / "PRUEBA_TRAZABILIDAD_RM.xlsx"
    ruta_log = c["rm"] / "log_consolidacion_rm.txt"

    try:
        consolidar(
            rutas,
            ruta_prueba=ruta_prueba,
            ruta_log=ruta_log,
            modo_reset=modo_reset,
            modo_prueba=modo_prueba,
        )
    except (ArchivoFaltante, PermissionError, ValueError) as e:
        print(f"\n✗ ERROR CRÍTICO: {e}\n")
        sys.exit(2)


if __name__ == "__main__":
    _main()
