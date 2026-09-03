"""
control_calidad.py — Control de calidad Zona RM (versión refactorizada)
Ambipar Group — Proyecto TRZ-APP-001, Etapa 1

QUÉ CAMBIÓ respecto de 02_control_calidad_rm.py:

    1. Ya no hay rutas escritas dentro del código. La función recibe un
       diccionario que dice dónde está cada archivo.

    2. Hay dos formas de armar ese diccionario:
         - rutas_desde_config()  → los archivos en su lugar de siempre en OneDrive
         - rutas_desde_carpeta() → todos los archivos juntos en una carpeta
       La segunda es la que va a usar la aplicación cuando la operadora suba
       los Excel, porque ahí van a llegar todos al mismo lugar.

    3. El código dejó de ejecutarse solo al abrir el archivo: ahora está
       dentro de funciones, que es lo que permite llamarlo desde otro programa.

    4. La función devuelve los resultados además de escribir el Excel.

QUÉ NO CAMBIÓ:

    Los nueve controles (C1 a C9) son exactamente los mismos. Ninguna regla
    de negocio fue modificada.

Se puede usar de dos formas:

    Desde otro programa:
        from control_calidad import controlar, rutas_desde_carpeta
        rutas = rutas_desde_carpeta("C:/carpeta/con/los/excel")
        resultado = controlar(rutas)

    Desde la terminal, igual que antes:
        python control_calidad.py
"""

import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


# ══════════════════════════════════════════════════════════════════
# CONSTANTES DE LA ZONA RM
# ══════════════════════════════════════════════════════════════════
AÑO_DESDE = 2026

# Movimientos que el consolidador reconoce en las fuentes que traen columna
# MOVIMIENTO (GIRI y Ecofibras San Bernardo). Cualquier otro valor hace que la
# fila se descarte silenciosamente al consolidar.
MOVIMIENTOS_VALIDOS_FUENTE = {"INGRESO", "TRASLADO"}

CRITICAS = ["Fecha", "Cliente", "RUT", "Generador", "TIPO", "Destino", "Peso neto"]

# Nombre esperado de cada archivo fuente. La clave es el nombre interno que
# usa el código; el valor son los nombres posibles del archivo, porque en
# OneDrive aparecen escritos de varias formas.
NOMBRES_ARCHIVO = {
    "bo":             ["BBDD BO SAN BERNARDO.xlsx"],
    "bo_valparaiso":  ["BBDD BO VALPARAÍSO.xlsx", "BBDD BO VALPARAISO.xlsx"],
    "ecofibras":      ["BBDD ECOFIBRAS SAN BERNARDO.xlsx"],
    "proveedores":    ["BBDD PROVEEDORES.xlsx"],
    "tradicionales":  ["BBDD TRADICIONALES.xlsx"],
    "giri":           ["OPERACIONES GIRI (INGRESOS Y EGRESOS).xlsx"],
    "sinader":        ["Clasificación_Residuos SINADER.xlsx", "Clasificacion_Residuos SINADER.xlsx"],
    "transportistas": ["Transportistas.xlsx"],
}

# Tabla de Excel que hay que leer dentro de cada archivo
TABLAS = {
    "bo":            "Tabla_operacion",
    "bo_valparaiso": "Tabla_operacion",
    "giri":          "INGRESO",
    "ecofibras":     "INGRESO",
    "proveedores":   "Tabla2",
    "tradicionales": "Analizado",
}


class ArchivoFaltante(Exception):
    """Se lanza cuando no se encuentra alguno de los archivos fuente."""


# ══════════════════════════════════════════════════════════════════
# REGISTRO DE MENSAJES
# ══════════════════════════════════════════════════════════════════
class Registro:
    """Guarda los mensajes además de mostrarlos, para que la app pueda usarlos."""

    def __init__(self, mostrar=True):
        self.lineas = []
        self.mostrar = mostrar

    def __call__(self, msg=""):
        self.lineas.append(str(msg))
        if self.mostrar:
            print(msg)

    def texto(self):
        return "\n".join(self.lineas)


# ══════════════════════════════════════════════════════════════════
# CÓMO ENCONTRAR LOS ARCHIVOS
# ══════════════════════════════════════════════════════════════════
def _clave(nombre):
    """Normaliza un nombre de archivo para poder compararlo con seguridad.

    SharePoint devuelve los nombres con las tildes descompuestas: la 'Í' puede
    venir como un solo carácter o como 'I' seguida de un acento invisible. Sin
    esta normalización, dos nombres que se ven idénticos no coinciden.
    También se ignoran mayúsculas, espacios, puntos y guiones.
    """
    n = unicodedata.normalize("NFKD", str(nombre))
    n = n.encode("ascii", "ignore").decode("ascii")
    n = n.upper()
    for c in " ._-()":
        n = n.replace(c, "")
    return n


def rutas_desde_config(base=None):
    """Arma las rutas usando la estructura de carpetas de siempre en OneDrive."""
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
        "sinader":        info / "Clasificación_Residuos SINADER.xlsx",
        "transportistas": info / "Transportistas.xlsx",
    }


def rutas_desde_carpeta(carpeta):
    """Arma las rutas buscando todos los archivos dentro de una sola carpeta.

    Es la forma que va a usar la aplicación: la operadora sube los ocho Excel
    y todos quedan juntos. Los nombres se comparan de forma tolerante, así que
    da lo mismo si vienen con tilde o sin tilde, con espacios de más, o en
    mayúsculas distintas.

    Si falta alguno, avisa cuáles faltan y cuáles sí encontró.
    """
    carpeta = Path(carpeta)
    if not carpeta.exists():
        raise ArchivoFaltante(f"La carpeta no existe:\n  {carpeta}")

    presentes = {_clave(f.name): f for f in carpeta.iterdir() if f.is_file()}

    rutas = {}
    faltantes = []
    for interno, posibles in NOMBRES_ARCHIVO.items():
        encontrado = None
        for nombre in posibles:
            encontrado = presentes.get(_clave(nombre))
            if encontrado:
                break
        if encontrado:
            rutas[interno] = encontrado
        else:
            faltantes.append(posibles[0])

    if faltantes:
        hay = "\n".join(f"    · {f.name}" for f in sorted(presentes.values()))
        falta = "\n".join(f"    · {f}" for f in faltantes)
        raise ArchivoFaltante(
            f"Faltan archivos en la carpeta:\n{falta}\n\n"
            f"  Archivos encontrados en {carpeta.name}:\n{hay if hay else '    (ninguno)'}"
        )

    return rutas


def verificar_rutas(rutas):
    """Comprueba que todos los archivos existan antes de empezar a leer.

    Vale la pena revisar todo junto y avisar de una sola vez, en vez de fallar
    en el quinto archivo después de haber esperado la lectura de los cuatro
    anteriores.
    """
    faltan = [f"{k}: {v}" for k, v in rutas.items() if not Path(v).exists()]
    if faltan:
        detalle = "\n".join(f"    · {f}" for f in faltan)
        raise ArchivoFaltante(f"No se encontraron estos archivos:\n{detalle}")


# ══════════════════════════════════════════════════════════════════
# FUNCIONES BÁSICAS
# ══════════════════════════════════════════════════════════════════
def leer_tabla(path, nombre_tabla, p):
    """Lee una tabla con nombre dentro de un Excel."""
    path = Path(path)
    p(f"  Leyendo {path.name} → tabla {nombre_tabla}...")
    wb = load_workbook(path, data_only=True)
    for sheet in wb.worksheets:
        for tabla in sheet.tables.values():
            if tabla.name == nombre_tabla:
                filas = [[c.value for c in fila] for fila in sheet[tabla.ref]]
                df    = pd.DataFrame(filas[1:], columns=filas[0])
                wb.close()
                df.columns = df.columns.astype(str).str.strip()
                return df.dropna(how="all").reset_index(drop=True)
    wb.close()
    raise ValueError(
        f"El archivo {path.name} no tiene una tabla llamada '{nombre_tabla}'.\n"
        f"  Revisa que la tabla no haya sido renombrada o eliminada en Excel."
    )


def normalizar(serie):
    return (
        serie.astype(str).str.strip().str.upper()
        .str.replace(r"\s+", " ", regex=True)
        .str.replace(r"\.$", "", regex=True)
        .str.replace(".", "", regex=False)
        .str.strip()
    )


def normalizar_rut(serie):
    return (
        serie.astype(str).str.strip()
        .str.replace(".", "", regex=False)
        .str.replace("-", "", regex=False)
        .str.upper()
    )


def tomar_columna(df, opciones, default=None):
    """Devuelve la primera columna que exista entre varios nombres posibles."""
    for op in opciones:
        if op in df.columns:
            return df[op]
    return pd.Series([default] * len(df), index=df.index)


def esta_vacio(serie):
    """Marca como vacío los nulos y los textos que representan vacío."""
    return (
        serie.isna()
        | serie.astype(str).str.strip().str.lower().isin(["", "nan", "none", "nat", "#n/a", "0"])
    )


# ══════════════════════════════════════════════════════════════════
# VISTA UNIFICADA DE FUENTES
# ══════════════════════════════════════════════════════════════════
# Los controles C7, C8 y C9 necesitan mirar todas las fuentes con los mismos
# nombres de columna. Los nombres de cada fuente se tomaron de las funciones
# procesar_* de 03_consolidar_rm.py.

def _mapas(dfs):
    return {
        "BO": {
            "df": dfs["bo"],
            "Fecha": ["Fecha"], "Cliente": ["Cliente"], "RUT": ["RUT"],
            "Generador": ["Generador"], "TIPO": ["TIPO"], "Destino": ["Destino"],
            "Peso neto": ["Peso neto"],
            "Observaciones": ["Observaciones", "Observación", "Obs"],
            "Movimiento": [], "mov_fijo": "TRASLADO", "necesita_destino": True,
        },
        "BO Valparaíso": {
            "df": dfs["bo_valparaiso"],
            "Fecha": ["Fecha"], "Cliente": ["Cliente"], "RUT": ["RUT"],
            "Generador": ["Generador"], "TIPO": ["TIPO"], "Destino": ["Destino"],
            "Peso neto": ["Peso neto [kg]", "Peso neto"],
            "Observaciones": [],
            "Movimiento": [], "mov_fijo": "TRASLADO", "necesita_destino": True,
        },
        "GIRI": {
            "df": dfs["giri"],
            "Fecha": ["FECHA"], "Cliente": ["CLIENTE"], "RUT": ["RUT CLIENTE"],
            "Generador": ["GENERADOR"], "TIPO": ["TIPO"], "Destino": ["DESTINO"],
            "Peso neto": ["PESO NETO KG"],
            "Observaciones": [],
            "Movimiento": ["MOVIMIENTO"], "mov_fijo": None,
            "necesita_destino": "solo_traslado",
        },
        "Ecofibras": {
            "df": dfs["ecofibras"],
            "Fecha": ["FECHA"], "Cliente": ["CLIENTE"], "RUT": ["RUT CLIENTE"],
            "Generador": ["GENERADOR"], "TIPO": ["TIPO"], "Destino": ["DESTINO"],
            "Peso neto": ["PESO NETO KG"],
            "Observaciones": [],
            "Movimiento": ["MOVIMIENTO"], "mov_fijo": None,
            "necesita_destino": "solo_traslado",
        },
        "Proveedores": {
            "df": dfs["proveedores"],
            "Fecha": ["Fecha"], "Cliente": ["Cliente"], "RUT": ["RUT CLIENTE"],
            "Generador": ["Generador"], "TIPO": ["TIPO"],
            "Destino": ["Destinatario", "Destino"],
            "Peso neto": ["Cantidad (Kg)"],
            "Observaciones": [],
            "Movimiento": [], "mov_fijo": "TRASLADO EXTERNO", "necesita_destino": True,
        },
        "Tradicionales": {
            "df": dfs["tradicionales"],
            "Fecha": ["Fec_salida"], "Cliente": ["Raz_social"], "RUT": ["Rut"],
            "Generador": ["Raz_social"],   # el consolidador copia el cliente
            "TIPO": [],                    # se fija en ASIMILABLE A DOMICILIARIO
            "Destino": ["Destino"],
            "Peso neto": ["Peso Final [kg]"],
            "Observaciones": [],
            "Movimiento": [], "mov_fijo": "TRASLADO", "necesita_destino": True,
        },
    }


def _vista_unificada(dfs, p):
    p("\nArmando vista unificada para los controles C7, C8 y C9...")

    partes = []
    for origen, mapa in _mapas(dfs).items():
        df_o = mapa["df"]
        out = pd.DataFrame(index=df_o.index)
        out["ORIGEN"] = origen

        for campo in ["Fecha", "Cliente", "RUT", "Generador", "TIPO", "Destino", "Peso neto", "Observaciones"]:
            out[campo] = tomar_columna(df_o, mapa[campo])

        if mapa["Movimiento"]:
            out["Movimiento"] = tomar_columna(df_o, mapa["Movimiento"])
        else:
            out["Movimiento"] = mapa["mov_fijo"]

        if origen == "Tradicionales":
            out["TIPO"] = "ASIMILABLE A DOMICILIARIO"

        out["MOV_NORM"] = out["Movimiento"].astype(str).str.strip().str.upper()
        out["Fecha"] = pd.to_datetime(out["Fecha"], errors="coerce")
        out["_desde"] = out["Fecha"].dt.year >= AÑO_DESDE
        out["_tiene_col_movimiento"] = bool(mapa["Movimiento"])
        out["_necesita_destino"] = (
            mapa["necesita_destino"] if mapa["necesita_destino"] != "solo_traslado" else None
        )
        out["_destino_solo_traslado"] = mapa["necesita_destino"] == "solo_traslado"

        partes.append(out)

    df_all = pd.concat(partes, ignore_index=True)
    df_all = df_all[df_all["_desde"]].copy()
    p(f"  Filas con fecha ≥ {AÑO_DESDE}: {len(df_all)}")
    return df_all


# ══════════════════════════════════════════════════════════════════
# CONTROLES
# ══════════════════════════════════════════════════════════════════
def _clientes_y_generadores(dfs, p):
    p("\nExtrayendo clientes...")
    specs = [
        (dfs["bo"],            "Cliente",    "RUT",         "BO"),
        (dfs["bo_valparaiso"], "Cliente",    "RUT",         "BO Valparaíso"),
        (dfs["giri"],          "CLIENTE",    "RUT CLIENTE", "GIRI"),
        (dfs["ecofibras"],     "CLIENTE",    "RUT CLIENTE", "Ecofibras"),
        (dfs["proveedores"],   "Cliente",    "RUT CLIENTE", "Proveedores"),
        (dfs["tradicionales"], "Raz_social", "Rut",         "Tradicionales"),
    ]

    pares = []
    for df, cn, cr, origen in specs:
        sub = df[[cn, cr]].copy()
        sub.columns = ["NOMBRE", "RUT"]
        sub["ORIGEN"]      = origen
        sub["NOMBRE_NORM"] = normalizar(sub["NOMBRE"])
        sub["RUT_NORM"]    = normalizar_rut(sub["RUT"])
        sub = sub[sub["NOMBRE_NORM"].str.lower() != "nan"]
        sub = sub[sub["RUT_NORM"].str.lower()    != "nan"]
        pares.append(sub)

    clientes = pd.concat(pares, ignore_index=True).drop_duplicates(
        subset=["NOMBRE_NORM", "RUT_NORM", "ORIGEN"]
    )
    p(f"  Total pares únicos: {len(clientes)}")

    p("\nExtrayendo generadores...")
    specs_gen = [
        (dfs["bo"],            "Generador", "BO"),
        (dfs["bo_valparaiso"], "Generador", "BO Valparaíso"),
        (dfs["giri"],          "GENERADOR", "GIRI"),
        (dfs["ecofibras"],     "GENERADOR", "Ecofibras"),
        (dfs["proveedores"],   "Generador", "Proveedores"),
    ]

    gens = []
    for df, col, origen in specs_gen:
        sub = df[[col]].copy()
        sub.columns = ["NOMBRE"]
        sub["ORIGEN"]      = origen
        sub["NOMBRE_NORM"] = normalizar(sub["NOMBRE"])
        sub = sub[sub["NOMBRE_NORM"].str.lower() != "nan"]
        gens.append(sub)

    generadores = pd.concat(gens, ignore_index=True).drop_duplicates(
        subset=["NOMBRE_NORM", "ORIGEN"]
    )
    p(f"  Total generadores únicos: {len(generadores)}")

    return clientes, generadores


def _c1(clientes, p):
    p("\nC1: Cliente mismo nombre normalizado, RUT distinto...")
    ruts = clientes.groupby("NOMBRE_NORM")["RUT_NORM"].nunique().reset_index(name="n")
    conf = ruts[ruts["n"] > 1]["NOMBRE_NORM"]
    if len(conf) > 0:
        c1 = (
            clientes[clientes["NOMBRE_NORM"].isin(conf)]
            [["NOMBRE", "NOMBRE_NORM", "RUT", "RUT_NORM", "ORIGEN"]]
            .drop_duplicates().sort_values(["NOMBRE_NORM", "RUT_NORM"])
        )
        c1.insert(0, "PROBLEMA", "Mismo nombre, RUT distinto")
    else:
        c1 = pd.DataFrame()
    p(f"  Conflictos: {len(c1)}")
    return c1


def _c2(clientes, p):
    p("C2: Mismo RUT, nombre distinto...")
    noms = clientes.groupby("RUT_NORM")["NOMBRE_NORM"].nunique().reset_index(name="n")
    conf2 = noms[noms["n"] > 1]["RUT_NORM"]
    if len(conf2) > 0:
        c2 = (
            clientes[clientes["RUT_NORM"].isin(conf2)]
            [["NOMBRE", "NOMBRE_NORM", "RUT", "RUT_NORM", "ORIGEN"]]
            .drop_duplicates().sort_values(["RUT_NORM", "NOMBRE_NORM"])
        )
        c2.insert(0, "PROBLEMA", "Mismo RUT, nombre distinto")
    else:
        c2 = pd.DataFrame()
    p(f"  Conflictos: {len(c2)}")
    return c2


def _c3(generadores, p):
    p("C3: Generadores con variantes de nombre...")
    vars_gen = generadores.groupby("NOMBRE_NORM")["NOMBRE"].nunique().reset_index(name="n")
    conf3 = vars_gen[vars_gen["n"] > 1]["NOMBRE_NORM"]
    if len(conf3) > 0:
        c3 = (
            generadores[generadores["NOMBRE_NORM"].isin(conf3)]
            [["NOMBRE", "NOMBRE_NORM", "ORIGEN"]]
            .drop_duplicates().sort_values(["NOMBRE_NORM", "NOMBRE"])
        )
        c3.insert(0, "PROBLEMA", "Generador con variantes de nombre")
    else:
        c3 = pd.DataFrame()
    p(f"  Conflictos: {len(c3)}")
    return c3


def _c4(clientes, p):
    p("C4: RUT Tradicionales vs otras bases...")
    otras = (
        clientes[clientes["ORIGEN"] != "Tradicionales"]
        .drop_duplicates(subset=["NOMBRE_NORM"])
        .set_index("NOMBRE_NORM")["RUT"]
        .to_dict()
    )

    trad_rut = clientes[clientes["ORIGEN"] == "Tradicionales"].copy()
    trad_rut["RUT_OTRA_BASE"] = trad_rut["NOMBRE_NORM"].map(otras)
    trad_rut["ESTADO"] = trad_rut["RUT_OTRA_BASE"].apply(
        lambda x: "✓ Match exacto" if pd.notna(x) else "Sin match"
    )
    c4 = trad_rut[["NOMBRE", "NOMBRE_NORM", "RUT", "RUT_OTRA_BASE", "ESTADO"]].drop_duplicates()
    con_match = (c4["ESTADO"] == "✓ Match exacto").sum()
    sin_match = (c4["ESTADO"] == "Sin match").sum()
    p(f"  Con match: {con_match}  |  Sin match: {sin_match}")
    return c4


def _c5(dfs, rutas, p):
    p("C5: TIPO sin código SINADER...")
    df_sin = pd.read_excel(rutas["sinader"], sheet_name="Clasificación Residuos", header=1)
    df_sin.columns = df_sin.columns.astype(str).str.strip()
    tipos_ok = set(df_sin["TIPO"].dropna().astype(str).str.strip())
    tipos_ok.add("ASIMILABLE A DOMICILIARIO")

    specs_tipo = [
        (dfs["bo"],            "TIPO", "BO"),
        (dfs["bo_valparaiso"], "TIPO", "BO Valparaíso"),
        (dfs["giri"],          "TIPO", "GIRI"),
        (dfs["ecofibras"],     "TIPO", "Ecofibras"),
        (dfs["proveedores"],   "TIPO", "Proveedores"),
    ]
    todos_tipos = []
    for df, col, origen in specs_tipo:
        if col in df.columns:
            sub = df[[col]].copy()
            sub.columns = ["TIPO"]
            sub["ORIGEN"] = origen
            todos_tipos.append(sub)

    df_tipos = pd.concat(todos_tipos, ignore_index=True)
    df_tipos["TIPO"] = df_tipos["TIPO"].astype(str).str.strip()
    df_tipos = df_tipos[df_tipos["TIPO"].str.lower() != "nan"].drop_duplicates()
    c5 = df_tipos[~df_tipos["TIPO"].isin(tipos_ok)].copy()
    if len(c5) > 0:
        c5.insert(0, "PROBLEMA", "TIPO sin código SINADER")
    else:
        c5 = pd.DataFrame()
    p(f"  Sin código: {len(c5)}")
    return c5


def _c6(dfs, rutas, p):
    p("C6: Gestores sin RUT en tabla Transportistas...")
    df_trans = pd.read_excel(rutas["transportistas"], sheet_name="Hoja1", header=0)
    df_trans.columns = df_trans.columns.astype(str).str.strip()
    ruts_ok = set(df_trans["Transportista"].dropna().astype(str).str.strip())

    specs_gest = [
        (dfs["giri"],        "GESTOR", "GIRI"),
        (dfs["ecofibras"],   "GESTOR", "Ecofibras"),
        (dfs["proveedores"], "Gestor", "Proveedores"),
    ]
    filas_gest = []
    for df, col, origen in specs_gest:
        if col in df.columns:
            for g in df[col].dropna().astype(str).str.strip().unique():
                if g and g.lower() != "nan" and g not in ruts_ok:
                    filas_gest.append({"GESTOR": g, "ORIGEN": origen})

    c6 = (
        pd.DataFrame(filas_gest).drop_duplicates(subset=["GESTOR"])
        if filas_gest else pd.DataFrame()
    )
    if len(c6) > 0:
        c6.insert(0, "PROBLEMA", "Gestor sin RUT en Transportistas")
    p(f"  Sin RUT: {len(c6)}")
    return c6


def _c7(df_all, p):
    """Vacíos críticos: campos sin los cuales la fila no es trazable.

    Incluye una columna MES para poder ver si el problema se concentra en un
    período puntual o si viene arrastrándose desde el inicio del año. El
    filtro de fondo (AÑO_DESDE) no cambia: se sigue revisando el año completo
    en una sola corrida, igual que lo consolida 03_consolidar_rm.py.
    """
    p("\nC7: Vacíos críticos...")
    filas_vacios = []
    for col in CRITICAS:
        vacio = esta_vacio(df_all[col])
        if vacio.any():
            sub = df_all.loc[vacio].copy()
            sub["MES"] = sub["Fecha"].dt.to_period("M").astype(str)
            por_origen_mes = sub.groupby(["ORIGEN", "MES"]).size().reset_index(name="CANTIDAD")
            for _, r in por_origen_mes.iterrows():
                filas_vacios.append({
                    "PROBLEMA": "Columna crítica vacía",
                    "COLUMNA":  col,
                    "ORIGEN":   r["ORIGEN"],
                    "MES":      r["MES"],
                    "CANTIDAD": int(r["CANTIDAD"]),
                })

    c7 = pd.DataFrame(filas_vacios)
    p(f"  Combinaciones fuente-columna-mes con vacíos: {len(c7)}")
    return c7


def _c8(df_all, p):
    """Movimientos no reconocidos: esas filas se descartan al consolidar."""
    p("C8: Movimientos a revisar...")
    con_columna_mov = df_all[df_all["_tiene_col_movimiento"]]
    mov_malo = ~con_columna_mov["MOV_NORM"].isin(MOVIMIENTOS_VALIDOS_FUENTE)

    c8 = con_columna_mov.loc[
        mov_malo, ["ORIGEN", "Fecha", "Cliente", "Movimiento", "Peso neto"]
    ].drop_duplicates().copy()

    if not c8.empty:
        c8.insert(0, "PROBLEMA", "Movimiento vacío o no reconocido (la fila se descarta al consolidar)")
    p(f"  Filas con movimiento no reconocido: {len(c8)}")
    return c8


def _tiene_observacion(serie):
    """True si la celda tiene texto real, no vacío ni un símbolo de relleno."""
    return ~esta_vacio(serie) & ~serie.astype(str).str.strip().isin(["-", "—", "s/o", "S/O"])


def _c9(df_all, p):
    """Destinos vacíos donde sí hacen falta: esas filas se descartan al consolidar.

    Se separan en dos grupos:
      - c9: sin destino y SIN ninguna observación que lo explique. Esta es la
        alerta fuerte, porque no hay ningún registro de por qué falta el dato.
      - c9_informativo: sin destino pero CON algo escrito en Observaciones
        (ej. "falta llenado de tolva", "no recibió cliente"). Son casos como
        fletes en falso: el servicio no se realizó, así que no corresponde
        exigirles destino. Quedan en una lista aparte para revisión, no como
        error de digitación.

    Por ahora la columna Observaciones solo existe en la fuente BO. En el
    resto de las fuentes la columna queda vacía y por lo tanto esas filas
    siempre caen en la alerta fuerte, igual que antes.

    Incluye la columna Mes para ver si el problema se concentra en un período.
    """
    p("C9: Destinos a revisar...")
    necesita = (
        df_all["_necesita_destino"].fillna(False).astype(bool)
        | (df_all["_destino_solo_traslado"] & df_all["MOV_NORM"].eq("TRASLADO"))
    )
    destino_vacio = esta_vacio(df_all["Destino"])
    con_observacion = _tiene_observacion(df_all["Observaciones"])

    columnas = ["ORIGEN", "Fecha", "Cliente", "Movimiento", "Destino", "Peso neto", "Observaciones"]

    base = df_all.loc[necesita & destino_vacio, columnas].copy()
    if not base.empty:
        base["MES"] = pd.to_datetime(base["Fecha"], errors="coerce").dt.to_period("M").astype(str)

    mask_obs = necesita & destino_vacio & con_observacion
    mask_sin_obs = necesita & destino_vacio & ~con_observacion

    c9 = df_all.loc[mask_sin_obs, columnas].copy()
    if not c9.empty:
        c9.insert(0, "PROBLEMA", "Destino vacío sin explicación (la fila se descarta al consolidar)")
        c9["MES"] = pd.to_datetime(c9["Fecha"], errors="coerce").dt.to_period("M").astype(str)

    c9_informativo = df_all.loc[mask_obs, columnas].copy()
    if not c9_informativo.empty:
        c9_informativo.insert(0, "PROBLEMA", "Destino vacío con observación operativa (ej. flete en falso)")
        c9_informativo["MES"] = pd.to_datetime(c9_informativo["Fecha"], errors="coerce").dt.to_period("M").astype(str)

    p(f"  Filas con destino vacío sin explicación: {len(c9)}")
    p(f"  Filas con destino vacío y observación operativa: {len(c9_informativo)}")
    return c9, c9_informativo


def _c9_resumen(c9, p):
    """Resumen de C9 por origen y mes: cuántas filas y cuántos kilos.

    Se calcula aparte para que quien abra el Excel vea de inmediato dónde se
    concentra el problema, sin tener que armar una tabla dinámica.
    """
    if c9.empty:
        return pd.DataFrame(columns=["ORIGEN", "MES", "FILAS", "PESO_TOTAL_KG"])

    c9 = c9.copy()
    # 'Peso neto' puede traer texto (comas, espacios, celdas mal tipeadas en el
    # Excel de origen). Se convierte a número antes de sumar; lo que no se
    # pueda convertir cuenta como 0 en el total, no rompe el cálculo.
    c9["Peso neto"] = pd.to_numeric(c9["Peso neto"], errors="coerce").fillna(0)

    resumen = (
        c9.groupby(["ORIGEN", "MES"])
        .agg(FILAS=("Cliente", "size"), PESO_TOTAL_KG=("Peso neto", "sum"))
        .reset_index()
        .sort_values(["ORIGEN", "MES"])
    )
    return resumen


# ══════════════════════════════════════════════════════════════════
# ESCRITURA DEL EXCEL DE SALIDA
# ══════════════════════════════════════════════════════════════════
def escribir_excel(res, ruta_salida, p):
    p(f"\nGuardando {Path(ruta_salida).name}...")

    c4 = res["c4"]
    resumen = pd.DataFrame({
        "Control": [
            "C1 — Cliente mismo nombre, RUT distinto",
            "C2 — Cliente mismo RUT, nombre distinto",
            "C3 — Generador con variantes",
            "C4 — RUT Tradicionales (informativo)",
            "C5 — TIPO sin código SINADER",
            "C6 — Gestores sin RUT",
            "C7 — Vacíos críticos",
            "C8 — Movimientos no reconocidos",
            "C9 — Destinos vacíos sin explicación",
            "C9b — Destinos vacíos con observación operativa (informativo)",
            "TOTAL",
        ],
        "Conflictos": [
            len(res["c1"]), len(res["c2"]), len(res["c3"]),
            int((c4["ESTADO"] == "Sin match").sum()) if not c4.empty else 0,
            len(res["c5"]), len(res["c6"]), len(res["c7"]),
            len(res["c8"]), len(res["c9"]), len(res["c9_informativo"]), res["total"],
        ],
        "Estado": [
            "✓ OK" if len(res["c1"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["c2"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["c3"]) == 0 else "⚠ Revisar",
            "Informativo",
            "✓ OK" if len(res["c5"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["c6"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["c7"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["c8"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["c9"]) == 0 else "⚠ Revisar",
            "Informativo",
            "✓ Listo" if res["total"] == 0 else "⚠ Corregir antes de consolidar",
        ],
    })

    def hoja(writer, df, nombre, msg_ok):
        sheet = df if not df.empty else pd.DataFrame({"resultado": [msg_ok]})
        sheet.to_excel(writer, sheet_name=nombre, index=False)

    info = pd.DataFrame({
        "Campo": ["Período de datos revisado", "Fecha de ejecución", "Filas analizadas", "Total conflictos"],
        "Valor": [res["periodo"], datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                  res["filas_analizadas"], res["total"]],
    })

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as w:
        info.to_excel(w, sheet_name="INFO", index=False)
        resumen.to_excel(w, sheet_name="RESUMEN", index=False)
        hoja(w, res["c1"], "C1_Cliente_RUT",       "Sin conflictos")
        hoja(w, res["c2"], "C2_RUT_nombre",        "Sin conflictos")
        hoja(w, res["c3"], "C3_Generadores",       "Sin variantes")
        hoja(w, res["c4"], "C4_RUT_Tradicionales", "Sin datos")
        hoja(w, res["c5"], "C5_TIPO_SINADER",      "Todos tienen código")
        hoja(w, res["c6"], "C6_Transportistas",    "Todos tienen RUT")
        hoja(w, res["c7"], "VACIOS_CRITICOS",      "Sin vacíos críticos")
        hoja(w, res["c8"], "MOVIMIENTOS_REVISAR",  "Movimientos OK")
        hoja(w, res["c9"], "DESTINOS_REVISAR",     "Sin destinos vacíos críticos")
        hoja(w, res["c9_informativo"], "DESTINOS_CON_OBSERVACION",
             "Sin casos con observación operativa")
        # DESTINOS_POR_MES y OBSERVACION_POR_MES ya no se escriben como hojas
        # aparte: son solo una agrupación por ORIGEN/MES de lo que ya está en
        # DESTINOS_REVISAR y DESTINOS_CON_OBSERVACION, y duplicaban el detalle.
        # res["c9_resumen"] y res["c9_informativo_resumen"] se siguen calculando
        # (quedan disponibles en el dict devuelto por controlar()) por si se
        # necesitan de nuevo, solo se dejó de volcarlos al Excel.

    p(f"✓ Guardado en: {ruta_salida}")


# ══════════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL — esta es la que llama la aplicación
# ══════════════════════════════════════════════════════════════════
def controlar(rutas, ruta_salida=None, mostrar=True):
    """Ejecuta el control de calidad de RM sobre las fuentes indicadas.

    Parámetros
    ----------
    rutas : dict
        Dónde está cada archivo. Se arma con rutas_desde_config() o con
        rutas_desde_carpeta().
    ruta_salida : str o Path, opcional
        Dónde escribir CONTROL_CALIDAD_RM.xlsx. Si se omite, no se escribe
        ningún archivo y los resultados solo se devuelven.
    mostrar : bool
        Si es True imprime el avance en pantalla.

    Devuelve
    --------
    dict con las claves:
        c1 ... c9   : un DataFrame por control
        total       : suma de conflictos (C4 no cuenta, es informativo)
        log         : todo el texto del avance
    """
    p = Registro(mostrar=mostrar)
    p("Iniciando control de calidad RM...")

    verificar_rutas(rutas)

    p("\nLeyendo bases...")
    dfs = {k: leer_tabla(rutas[k], tabla, p) for k, tabla in TABLAS.items()}

    p(f"  BO:            {len(dfs['bo'])} filas")
    p(f"  BO Valparaíso: {len(dfs['bo_valparaiso'])} filas")
    p(f"  GIRI:          {len(dfs['giri'])} filas")
    p(f"  Ecofibras:     {len(dfs['ecofibras'])} filas")
    p(f"  Proveedores:   {len(dfs['proveedores'])} filas")
    p(f"  Tradicionales: {len(dfs['tradicionales'])} filas")

    clientes, generadores = _clientes_y_generadores(dfs, p)
    df_all = _vista_unificada(dfs, p)

    # Período de datos analizados: informativo, para que quede claro qué mes(es)
    # cubre esta corrida. No es un filtro ni cambia ningún cálculo de los
    # controles C1 a C9 — solo se muestra en la hoja INFO del Excel de salida.
    fechas_validas = pd.to_datetime(df_all.get("Fecha"), errors="coerce").dropna()
    if not fechas_validas.empty:
        periodo = f"{fechas_validas.min().strftime('%m/%Y')} a {fechas_validas.max().strftime('%m/%Y')}"
    else:
        periodo = "No se pudo determinar (columna Fecha vacía o ausente)"

    res = {
        "c1": _c1(clientes, p),
        "c2": _c2(clientes, p),
        "c3": _c3(generadores, p),
        "c4": _c4(clientes, p),
        "c5": _c5(dfs, rutas, p),
        "c6": _c6(dfs, rutas, p),
        "c7": _c7(df_all, p),
        "c8": _c8(df_all, p),
        "c9": None,  # se completa abajo, porque _c9 ahora devuelve dos listas
    }
    res["c9"], res["c9_informativo"] = _c9(df_all, p)
    res["c9_resumen"] = _c9_resumen(res["c9"], p)
    res["c9_informativo_resumen"] = _c9_resumen(res["c9_informativo"], p)

    # C4 no suma: es informativo, no un conflicto que impida consolidar.
    res["total"] = sum(len(res[k]) for k in ["c1", "c2", "c3", "c5", "c6", "c7", "c8", "c9"])
    res["periodo"] = periodo
    res["filas_analizadas"] = len(df_all)

    if ruta_salida:
        # Se agrega fecha y hora al nombre para que cada corrida quede aparte
        # y no pise la anterior. Esto es temporal para la Etapa 1: en la
        # Etapa 2 el destino del archivo lo decide la aplicación, no una
        # carpeta fija en este computador.
        ruta_salida = Path(ruta_salida)
        sello = datetime.now().strftime("%Y-%m-%d_%H%M")
        ruta_salida = ruta_salida.with_name(f"{ruta_salida.stem}_{sello}{ruta_salida.suffix}")
        escribir_excel(res, ruta_salida, p)

    p()
    if res["total"] == 0:
        p("✓ Sin problemas — se puede consolidar")
    else:
        p(f"⚠ {res['total']} conflictos encontrados — revisa CONTROL_CALIDAD_RM.xlsx")

    res["log"] = p.texto()
    return res


# ══════════════════════════════════════════════════════════════════
# USO DESDE LA TERMINAL
# ══════════════════════════════════════════════════════════════════
def _main():
    """Permite ejecutar este archivo desde la terminal.

        python control_calidad.py
            → lee los archivos desde OneDrive, como siempre

        python control_calidad.py C:/carpeta/con/los/excel
            → lee todos los archivos desde una sola carpeta
    """
    argumentos = [a for a in sys.argv[1:] if not a.startswith("--")]

    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    try:
        from config import carpetas, ConfiguracionFaltante
    except ImportError:
        print("\n✗ No se encontró config.py. Debe estar en la carpeta src/.\n")
        sys.exit(2)

    try:
        if argumentos:
            rutas = rutas_desde_carpeta(argumentos[0])
            salida = Path(argumentos[0]) / "CONTROL_CALIDAD_RM.xlsx"
        else:
            rutas = rutas_desde_config()
            salida = carpetas()["rm"] / "CONTROL_CALIDAD_RM.xlsx"
    except ConfiguracionFaltante as e:
        print(f"\n✗ {e}\n")
        sys.exit(2)
    except ArchivoFaltante as e:
        print(f"\n✗ {e}\n")
        sys.exit(2)

    try:
        res = controlar(rutas, ruta_salida=salida)
    except ArchivoFaltante as e:
        print(f"\n✗ {e}\n")
        sys.exit(2)
    except ValueError as e:
        print(f"\n✗ {e}\n")
        sys.exit(2)

    sys.exit(0 if res["total"] == 0 else 1)


if __name__ == "__main__":
    _main()
