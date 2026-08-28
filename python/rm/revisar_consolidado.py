"""
revisar_consolidado.py — Revisión del consolidado RM (versión refactorizada)
Ambipar Group — Proyecto TRZ-APP-001, Etapa 1

QUÉ CAMBIÓ respecto de 04_revisar_consolidado_rm.py:

    1. Las rutas ya no están escritas dentro del código: se reciben como
       parámetro. Esto permite revisar un archivo que esté en OneDrive, en una
       carpeta temporal, o en cualquier otro lugar.

    2. Ya no se usa '--prueba' desde la línea de comandos: es un parámetro
       normal de la función.

    3. La función devuelve los resultados además de escribir el Excel, para
       que una aplicación pueda mostrar las alertas en pantalla sin abrir el
       archivo.

    4. Los mensajes se pueden capturar en vez de solo imprimirse.

QUÉ NO CAMBIÓ:

    La lógica de las diez validaciones (V0 a V9) es exactamente la misma,
    línea por línea. Ninguna regla de negocio fue modificada.

Se puede usar de dos formas:

    Desde otro programa:
        from revisar_consolidado import revisar
        resultado = revisar("C:/ruta/PRUEBA_TRAZABILIDAD_RM.xlsx")
        print(resultado["total_alertas"])

    Desde la terminal, igual que antes:
        python revisar_consolidado.py --prueba
        python revisar_consolidado.py C:/ruta/cualquier_archivo.xlsx
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


# ══════════════════════════════════════════════════════════════════
# CONSTANTES DE LA ZONA RM
# ══════════════════════════════════════════════════════════════════
# Estas sí son propias de la zona y viven acá, no en el archivo de
# configuración: son reglas de negocio, no rutas.

HOJA_DESTINO = "TRAZABILIDAD"
AÑO_DESDE    = 2026

COLUMNAS_FINALES = [
    "Fecha", "Mes", "Cliente", "RUT", "Gestor", "Contrato", "Generador",
    "Transportista", "Rut transportista", "Patente de Camión", "ID", "Origen ID",
    "Peso neto (kg)", "Destino", "Comuna Destino", "RUT DESTINATARIO",
    "TIPO", "CÓDIGOS SINADER", "Movimiento", "Movimiento interempresa",
    "CÓDIGO ESTABLECIMIENTO SINADER", "CÓDIGO DE TRATAMIENTO SINADER",
    "Región", "Destino inferido",
]

COLUMNAS_CRITICAS = [
    "Fecha", "Cliente", "RUT", "Gestor", "Generador",
    "Transportista", "Rut transportista", "Peso neto (kg)",
    "Destino", "Comuna Destino", "RUT DESTINATARIO",
    "TIPO", "CÓDIGOS SINADER", "Movimiento", "Movimiento interempresa",
    "CÓDIGO ESTABLECIMIENTO SINADER", "CÓDIGO DE TRATAMIENTO SINADER",
    "Región",
]

REGIONES_VALIDAS = {
    "Región Metropolitana",
    "Región de Valparaíso",   # BO Valparaíso
}

MOVIMIENTOS_VALIDOS = {"Ingreso", "Traslado", "Traslado externo"}

# El consolidado ya no publica "Ticket de pesaje"; su lugar lo toma "ID".
CLAVE_DUPLICADO = ["Fecha", "Cliente", "ID", "TIPO", "Destino", "Peso neto (kg)", "Contrato"]


# ══════════════════════════════════════════════════════════════════
# REGISTRO DE MENSAJES
# ══════════════════════════════════════════════════════════════════
class Registro:
    """Guarda los mensajes además de mostrarlos.

    Antes los mensajes solo se imprimían en la terminal. Ahora también se
    acumulan, para que una aplicación pueda mostrarlos en pantalla.
    """

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
# V0: ESTRUCTURA DE COLUMNAS
# ══════════════════════════════════════════════════════════════════
def validar_columnas(df, p):
    p("\n── V0: Estructura de columnas ──")

    filas = []
    faltantes = [c for c in COLUMNAS_FINALES if c not in df.columns]
    extras    = [c for c in df.columns if c not in COLUMNAS_FINALES]

    for c in faltantes:
        filas.append({"PROBLEMA": "Columna faltante", "COLUMNA": c})
        p(f"  ⚠ Falta columna: {c}")

    for c in extras:
        filas.append({"PROBLEMA": "Columna extra", "COLUMNA": c})
        p(f"  ⚠ Columna extra: {c}")

    if not filas:
        p("  ✓ Columnas correctas")

    return pd.DataFrame(filas) if filas else pd.DataFrame(columns=["PROBLEMA", "COLUMNA"])


# ══════════════════════════════════════════════════════════════════
# V1: COLUMNAS SIN VACÍOS
# ══════════════════════════════════════════════════════════════════
def validar_vacios(df, p):
    p("\n── V1: Columnas sin vacíos ──")
    filas_malas = []
    for col in COLUMNAS_CRITICAS:
        if col not in df.columns:
            filas_malas.append({
                "PROBLEMA": f"Columna '{col}' no existe en el archivo",
                "COLUMNA":  col,
                "CANTIDAD": "—",
                "PCT":      "—",
            })
            continue
        n_vacios = df[col].isna().sum() + (df[col].astype(str).str.lower() == "nan").sum()
        n_vacios = int(n_vacios)
        if n_vacios > 0:
            pct = round(n_vacios / len(df) * 100, 1)
            filas_malas.append({
                "PROBLEMA": "Columna con valores vacíos",
                "COLUMNA":  col,
                "CANTIDAD": n_vacios,
                "PCT":      f"{pct}%",
            })
            p(f"  ⚠ {col:<35} {n_vacios} vacíos ({pct}%)")
        else:
            p(f"  ✓ {col:<35} completa")

    det = pd.DataFrame(filas_malas) if filas_malas else pd.DataFrame(
        columns=["PROBLEMA", "COLUMNA", "CANTIDAD", "PCT"]
    )
    p(f"  Columnas con problemas: {len(det)}")
    return det


# ══════════════════════════════════════════════════════════════════
# V2: RUT EN FORMATO CORRECTO
# ══════════════════════════════════════════════════════════════════
def validar_rut(df, p):
    p("\n── V2: RUT en formato correcto ──")
    filas = []
    for _, row in df.iterrows():
        rut = str(row.get("RUT", "")).strip()
        cliente = row.get("Cliente", "")
        if rut and rut.lower() not in ("nan", "none", ""):
            if "-" not in rut:
                filas.append({
                    "PROBLEMA": "RUT sin dígito verificador",
                    "CLIENTE":  cliente,
                    "RUT":      rut,
                })

    det = pd.DataFrame(filas).drop_duplicates() if filas else pd.DataFrame(
        columns=["PROBLEMA", "CLIENTE", "RUT"]
    )
    p(f"  RUT sin formato correcto: {len(det)}")
    return det


# ══════════════════════════════════════════════════════════════════
# V3: DESTINOS INFERIDOS
# ══════════════════════════════════════════════════════════════════
def validar_destinos_inferidos(df, p):
    p("\n── V3: Destinos inferidos ──")
    if "Destino inferido" not in df.columns:
        p("  Columna 'Destino inferido' no existe")
        return pd.DataFrame()

    inferidos = df[df["Destino inferido"].astype(str) != "No"].copy()
    if len(inferidos) > 0:
        det = inferidos[["Fecha", "Cliente", "Destino", "Destino inferido"]].drop_duplicates()
        det.insert(0, "PROBLEMA", "Destino inferido automáticamente")
        p(f"  ⚠ Filas con destino inferido: {len(det)}")
        p(f"    (revisar la columna 'Destino inferido' en el consolidado)")
    else:
        det = pd.DataFrame(columns=["PROBLEMA", "Fecha", "Cliente", "Destino", "Destino inferido"])
        p(f"  ✓ Sin destinos inferidos")
    return det


# ══════════════════════════════════════════════════════════════════
# V4: DUPLICADOS
# ══════════════════════════════════════════════════════════════════
def validar_duplicados(df, p):
    p("\n── V4: Filas duplicadas ──")
    if not all(c in df.columns for c in CLAVE_DUPLICADO):
        p("  ⚠ Faltan columnas clave para detectar duplicados")
        return pd.DataFrame()

    df_temp = df.copy()
    for c in CLAVE_DUPLICADO:
        df_temp[c] = df_temp[c].astype(str).str.strip()

    dup_mask = df_temp.duplicated(subset=CLAVE_DUPLICADO, keep=False)
    duplicados = df_temp[dup_mask].copy()

    if len(duplicados) > 0:
        det = duplicados[CLAVE_DUPLICADO].sort_values(CLAVE_DUPLICADO)
        det.insert(0, "PROBLEMA", "Fila duplicada (misma clave)")
        p(f"  ⚠ Filas duplicadas: {len(det)}")
    else:
        det = pd.DataFrame(columns=["PROBLEMA"] + CLAVE_DUPLICADO)
        p(f"  ✓ Sin duplicados")

    return det


# ══════════════════════════════════════════════════════════════════
# V5: FECHAS FUERA DE RANGO
# ══════════════════════════════════════════════════════════════════
def validar_fechas(df, p):
    p("\n── V5: Fechas fuera de rango ──")
    df_temp = df.copy()
    df_temp["Fecha"] = pd.to_datetime(df_temp["Fecha"], errors="coerce")

    fechas_nulas = df_temp[df_temp["Fecha"].isna()]
    fechas_antiguas = df_temp[
        (df_temp["Fecha"].notna()) & (df_temp["Fecha"].dt.year < AÑO_DESDE)
    ]
    fechas_futuras = df_temp[df_temp["Fecha"] > datetime.now()]

    filas = []
    if len(fechas_nulas) > 0:
        filas.append({
            "PROBLEMA":  "Fechas inválidas (NaT)",
            "CANTIDAD":  len(fechas_nulas),
            "EJEMPLO":   "—",
        })
    if len(fechas_antiguas) > 0:
        filas.append({
            "PROBLEMA":  f"Fechas anteriores a {AÑO_DESDE}",
            "CANTIDAD":  len(fechas_antiguas),
            "EJEMPLO":   str(fechas_antiguas["Fecha"].min().date()),
        })
    if len(fechas_futuras) > 0:
        filas.append({
            "PROBLEMA":  "Fechas en el futuro",
            "CANTIDAD":  len(fechas_futuras),
            "EJEMPLO":   str(fechas_futuras["Fecha"].max().date()),
        })

    det = pd.DataFrame(filas) if filas else pd.DataFrame(
        columns=["PROBLEMA", "CANTIDAD", "EJEMPLO"]
    )
    p(f"  Problemas de fecha: {len(det)}")
    return det


# ══════════════════════════════════════════════════════════════════
# V6: CONTEO POR MES Y DESTINO
# ══════════════════════════════════════════════════════════════════
def conteo_por_mes(df, p):
    p("\n── V6: Conteo por mes y destino ──")
    df_temp = df.copy()
    df_temp["Fecha"] = pd.to_datetime(df_temp["Fecha"], errors="coerce")
    df_temp["Año-Mes"] = df_temp["Fecha"].dt.to_period("M").astype(str)
    df_temp["Peso neto (kg)"] = pd.to_numeric(df_temp["Peso neto (kg)"], errors="coerce")

    conteo = (
        df_temp.groupby(["Año-Mes", "Destino"])
        .agg(
            Filas=("Fecha", "count"),
            Peso_total_kg=("Peso neto (kg)", "sum"),
        )
        .reset_index()
        .sort_values(["Año-Mes", "Destino"])
    )
    p(f"  Combinaciones mes-destino: {len(conteo)}")
    return conteo


# ══════════════════════════════════════════════════════════════════
# V7: PESOS VÁLIDOS
# ══════════════════════════════════════════════════════════════════
def validar_pesos(df, p):
    p("\n── V7: Pesos válidos ──")
    df_temp = df.copy()
    df_temp["Peso neto (kg)"] = pd.to_numeric(df_temp["Peso neto (kg)"], errors="coerce")

    invalidos = df_temp[
        (df_temp["Peso neto (kg)"].isna()) | (df_temp["Peso neto (kg)"] <= 0)
    ]

    if len(invalidos) > 0:
        cols = [c for c in ["Fecha", "Cliente", "Destino", "TIPO", "Peso neto (kg)"] if c in invalidos.columns]
        det = invalidos[cols].copy()
        det.insert(0, "PROBLEMA", "Peso nulo, cero o negativo")
        p(f"  ⚠ Filas con peso inválido: {len(det)}")
    else:
        det = pd.DataFrame(columns=["PROBLEMA", "Fecha", "Cliente", "Destino", "TIPO", "Peso neto (kg)"])
        p("  ✓ Todos los pesos válidos")

    return det


# ══════════════════════════════════════════════════════════════════
# V8: REGIONES VÁLIDAS
# ══════════════════════════════════════════════════════════════════
def validar_regiones(df, p):
    p("\n── V8: Regiones válidas ──")

    if "Región" not in df.columns:
        p("  ⚠ No existe columna Región")
        return pd.DataFrame({"PROBLEMA": ["No existe columna Región"]})

    bad = df[~df["Región"].astype(str).isin(REGIONES_VALIDAS)][["Fecha", "Cliente", "Región"]].drop_duplicates()

    if not bad.empty:
        bad.insert(0, "PROBLEMA", "Región no válida para RM")
        p(f"  ⚠ Regiones inválidas: {len(bad)}")
        return bad

    p("  ✓ Regiones válidas")
    return pd.DataFrame(columns=["PROBLEMA", "Fecha", "Cliente", "Región"])


# ══════════════════════════════════════════════════════════════════
# V9: MOVIMIENTOS VÁLIDOS
# ══════════════════════════════════════════════════════════════════
def validar_movimientos(df, p):
    p("\n── V9: Movimientos válidos ──")

    if "Movimiento" not in df.columns:
        p("  ⚠ No existe columna Movimiento")
        return pd.DataFrame({"PROBLEMA": ["No existe columna Movimiento"]})

    bad = df[~df["Movimiento"].astype(str).isin(MOVIMIENTOS_VALIDOS)][["Fecha", "Cliente", "Movimiento"]].drop_duplicates()

    if not bad.empty:
        bad.insert(0, "PROBLEMA", "Movimiento no válido para RM")
        p(f"  ⚠ Movimientos inválidos: {len(bad)}")
        return bad

    p("  ✓ Movimientos válidos")
    return pd.DataFrame(columns=["PROBLEMA", "Fecha", "Cliente", "Movimiento"])


# ══════════════════════════════════════════════════════════════════
# ESCRITURA DEL EXCEL DE SALIDA
# ══════════════════════════════════════════════════════════════════
def escribir_excel(res, ruta_salida, nombre_archivo, p):
    """Escribe REVISION_CONSOLIDADO_RM.xlsx con una hoja por validación."""
    p(f"\n── Guardando {Path(ruta_salida).name} ──")

    total = res["total_alertas"]

    resumen = pd.DataFrame({
        "Validación": [
            "V0 — Estructura columnas",
            "V1 — Columnas con vacíos",
            "V2 — RUT sin guión",
            "V3 — Destinos inferidos (revisar)",
            "V4 — Filas duplicadas",
            "V5 — Fechas fuera de rango",
            "V7 — Pesos inválidos",
            "V8 — Regiones válidas",
            "V9 — Movimientos válidos",
            "TOTAL",
        ],
        "Alertas": [
            len(res["v0"]), len(res["v1"]), len(res["v2"]), len(res["v3"]),
            len(res["v4"]), len(res["v5"]), len(res["v7"]), len(res["v8"]),
            len(res["v9"]), total,
        ],
        "Estado": [
            "✓ OK" if len(res["v0"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v1"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v2"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v3"]) == 0 else "⚠ Revisar manualmente",
            "✓ OK" if len(res["v4"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v5"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v7"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v8"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v9"]) == 0 else "⚠ Revisar",
            "✓ Consolidado correcto" if total == 0 else "⚠ Hay alertas",
        ],
    })

    def hoja(w, df, nombre, msg_ok):
        out = df if not df.empty else pd.DataFrame({"resultado": [msg_ok]})
        out.to_excel(w, sheet_name=nombre, index=False)

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as w:
        info = pd.DataFrame({
            "Campo": ["Archivo revisado", "Fecha de revisión", "Alertas totales"],
            "Valor": [nombre_archivo, datetime.now().strftime("%d/%m/%Y %H:%M:%S"), total],
        })
        info.to_excel(w, sheet_name="INFO", index=False)

        resumen.to_excel(w, sheet_name="RESUMEN", index=False)
        hoja(w, res["v0"], "V0_Columnas",           "Columnas OK")
        hoja(w, res["v1"], "V1_Vacios",             "Todas las columnas completas")
        hoja(w, res["v2"], "V2_RUT_sin_guion",      "Todos los RUT con guión")
        hoja(w, res["v3"], "V3_Destinos_inferidos", "Sin destinos inferidos")
        hoja(w, res["v4"], "V4_Duplicados",         "Sin duplicados")
        hoja(w, res["v5"], "V5_Fechas",             "Sin fechas problemáticas")
        res["v6"].to_excel(w, sheet_name="V6_Conteo_mensual", index=False)
        hoja(w, res["v7"], "V7_Pesos",              "Todos los pesos válidos")
        hoja(w, res["v8"], "V8_Regiones",           "Regiones OK")
        hoja(w, res["v9"], "V9_Movimientos",        "Movimientos OK")

    p(f"  ✓ Guardado en: {ruta_salida}")


# ══════════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL — esta es la que llama la aplicación
# ══════════════════════════════════════════════════════════════════
def revisar(ruta_consolidado, ruta_salida=None, hoja=HOJA_DESTINO, mostrar=True):
    """Revisa un consolidado de RM y devuelve las alertas encontradas.

    Parámetros
    ----------
    ruta_consolidado : str o Path
        Archivo Excel a revisar. Puede estar en cualquier carpeta.
    ruta_salida : str o Path, opcional
        Dónde escribir REVISION_CONSOLIDADO_RM.xlsx. Si se omite, no se
        escribe ningún archivo y los resultados solo se devuelven.
    hoja : str
        Hoja del archivo a revisar. Por defecto TRAZABILIDAD.
    mostrar : bool
        Si es True imprime el avance en pantalla. La aplicación lo pone en
        False y usa el texto devuelto en 'log'.

    Devuelve
    --------
    dict con las claves:
        v0 ... v9        : un DataFrame por validación
        total_alertas    : número total (V6 no cuenta, es informativo)
        filas, columnas  : tamaño del archivo revisado
        log              : todo el texto del avance
        archivo          : nombre del archivo revisado
    """
    p = Registro(mostrar=mostrar)
    ruta = Path(ruta_consolidado)

    p("\n" + "═" * 60)
    p("  REVISIÓN DEL CONSOLIDADO — RM")
    p(f"  Archivo: {ruta.name}")
    p(f"  {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    p("═" * 60)

    if not ruta.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo a revisar:\n  {ruta}\n"
            f"  Revisa que la ruta esté bien escrita y que el archivo exista."
        )

    try:
        df = pd.read_excel(ruta, sheet_name=hoja)
    except ValueError:
        disponibles = pd.ExcelFile(ruta).sheet_names
        raise ValueError(
            f"El archivo {ruta.name} no tiene la hoja '{hoja}'.\n"
            f"  Hojas disponibles: {', '.join(disponibles)}"
        )

    p(f"\nLeyendo {ruta.name}...")
    p(f"  Total filas: {len(df)}")
    p(f"  Total columnas: {len(df.columns)}")

    res = {
        "v0": validar_columnas(df, p),
        "v1": validar_vacios(df, p),
        "v2": validar_rut(df, p),
        "v3": validar_destinos_inferidos(df, p),
        "v4": validar_duplicados(df, p),
        "v5": validar_fechas(df, p),
        "v6": conteo_por_mes(df, p),
        "v7": validar_pesos(df, p),
        "v8": validar_regiones(df, p),
        "v9": validar_movimientos(df, p),
    }

    # V6 no suma: es un conteo informativo, no un problema.
    res["total_alertas"] = sum(
        len(res[k]) for k in ["v0", "v1", "v2", "v3", "v4", "v5", "v7", "v8", "v9"]
    )
    res["filas"] = len(df)
    res["columnas"] = len(df.columns)
    res["archivo"] = ruta.name

    if ruta_salida:
        escribir_excel(res, ruta_salida, ruta.name, p)

    p("\n" + "═" * 60)
    if res["total_alertas"] == 0:
        p("  ✓ Consolidado sin alertas — todo correcto")
    else:
        p(f"  ⚠ {res['total_alertas']} alertas encontradas")
    p("═" * 60 + "\n")

    res["log"] = p.texto()
    return res


# ══════════════════════════════════════════════════════════════════
# USO DESDE LA TERMINAL — se conserva para poder seguir trabajando igual
# ══════════════════════════════════════════════════════════════════
def _main():
    """Permite ejecutar este archivo desde la terminal.

        python revisar_consolidado.py
            → revisa PRUEBA_TRAZABILIDAD_RM.xlsx en consolidador_trazabilidad/RM

        python revisar_consolidado.py C:/ruta/otro_archivo.xlsx
            → revisa el archivo que se le indique

    La aplicación NO usa esta función: llama directamente a revisar().
    """
    argumentos = [a for a in sys.argv[1:] if not a.startswith("--")]

    if argumentos:
        # Se indicó el archivo a mano: no hace falta leer la configuración.
        ruta = Path(argumentos[0])
        salida = ruta.parent / "REVISION_CONSOLIDADO_RM.xlsx"
    else:
        # Sin argumentos se deduce la ruta desde la configuración local.
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
        try:
            from config import archivo_prueba, ConfiguracionFaltante
        except ImportError:
            print("\n✗ No se encontró config.py. Debe estar en la carpeta src/.\n")
            sys.exit(2)

        try:
            ruta = archivo_prueba("rm")
        except ConfiguracionFaltante as e:
            print(f"\n✗ {e}\n")
            sys.exit(2)
        except Exception as e:
            print(f"\n✗ No se pudo leer la configuración: {e}\n")
            sys.exit(2)

        salida = ruta.parent / "REVISION_CONSOLIDADO_RM.xlsx"

    try:
        res = revisar(ruta, ruta_salida=salida)
    except (FileNotFoundError, ValueError) as e:
        print(f"\n✗ {e}\n")
        sys.exit(2)

    sys.exit(0 if res["total_alertas"] == 0 else 1)


if __name__ == "__main__":
    _main()
