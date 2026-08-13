"""
revisar_consolidado.py — Revisión del consolidado Norte (versión refactorizada)
Ambipar Group — Proyecto TRZ-APP-001, Etapa 1

Mismo patrón que src/rm y src/sur: la lógica de las diez validaciones (V0 a
V9) es exactamente la misma que en 03_revisar_consolidado_norte.py, línea
por línea. Lo único que cambia es que las rutas se reciben como parámetro y
la función devuelve los resultados además de escribir el Excel.

NOTA — Hallazgo 6 (11-08-2026): el script original tiene tres mensajes que
dicen "SUR" en vez de "NORTE" (quedaron de cuando se copió el archivo de esa
zona): en V3, en V9, y en el mensaje final al guardar. Es solo texto que se
imprime en pantalla y en el log; no afecta ningún cálculo ni ninguna regla
de negocio. Se preserva tal cual está en el original, porque corregir texto
no es parte de esta etapa — se documenta acá para que la dueña del proceso
decida si corregirlo aparte.

Se puede usar de dos formas:

    Desde otro programa:
        from revisar_consolidado import revisar
        resultado = revisar("C:/ruta/PRUEBA_TRAZABILIDAD_NORTE.xlsx")

    Desde la terminal, igual que antes:
        python revisar_consolidado.py
        python revisar_consolidado.py C:/ruta/cualquier_archivo.xlsx
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


# ══════════════════════════════════════════════════════════════════
# CONSTANTES DE LA ZONA NORTE
# ══════════════════════════════════════════════════════════════════
AÑO_DESDE = 2026
HOJA_DESTINO = "TRAZABILIDAD"

COLUMNAS_FINALES = [
    "Fecha", "Mes", "Cliente", "RUT", "Gestor", "Contrato", "Generador",
    "Transportista", "Rut transportista", "Patente de Camión", "Ticket de pesaje",
    "Peso neto (kg)", "Destino", "Comuna Destino", "RUT DESTINATARIO",
    "TIPO", "CÓDIGOS SINADER", "Movimiento", "Movimiento interempresa",
    "CÓDIGO ESTABLECIMIENTO SINADER", "CÓDIGO DE TRATAMIENTO SINADER",
    "Región", "Destino inferido",
]

COLUMNAS_CRITICAS = [
    "Fecha", "Cliente", "RUT", "Gestor", "Generador",
    "Transportista", "Rut transportista", "Peso neto (kg)",
    "Destino", "Comuna Destino", "RUT DESTINATARIO",
    "TIPO", "CÓDIGOS SINADER", "Movimiento",
    "CÓDIGO ESTABLECIMIENTO SINADER", "CÓDIGO DE TRATAMIENTO SINADER",
    "Región",
]

REGIONES_VALIDAS = {
    "Región de Tarapacá",
    "Región de Arica y Parinacota",
}

MOVIMIENTOS_VALIDOS = {"Ingreso", "Traslado"}

CLAVE_DUPLICADO = ["Fecha", "Cliente", "Ticket de pesaje", "TIPO", "Destino", "Peso neto (kg)", "Contrato"]


# ══════════════════════════════════════════════════════════════════
# REGISTRO DE MENSAJES
# ══════════════════════════════════════════════════════════════════
class Registro:
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
# VALIDACIONES — idénticas a 03_revisar_consolidado_norte.py
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


def validar_vacios(df, p):
    p("\n── V1: Columnas críticas sin vacíos ──")
    filas = []
    for col in COLUMNAS_CRITICAS:
        if col not in df.columns:
            filas.append({"PROBLEMA": "Columna no existe", "COLUMNA": col, "CANTIDAD": "—", "PCT": "—"})
            continue
        serie  = df[col]
        vacios = serie.isna() | serie.astype(str).str.strip().str.lower().isin(["", "nan", "none", "nat", "#n/a"])
        n = int(vacios.sum())
        if n > 0:
            pct = round(n / max(len(df), 1) * 100, 1)
            filas.append({"PROBLEMA": "Columna con valores vacíos", "COLUMNA": col, "CANTIDAD": n, "PCT": f"{pct}%"})
            p(f"  ⚠ {col:<42} {n} vacíos ({pct}%)")
        else:
            p(f"  ✓ {col:<42} completa")
    return pd.DataFrame(filas) if filas else pd.DataFrame(columns=["PROBLEMA", "COLUMNA", "CANTIDAD", "PCT"])


def validar_rut(df, p):
    p("\n── V2: RUT con guión ──")
    filas = []
    for _, row in df.iterrows():
        rut = str(row.get("RUT", "")).strip()
        if rut and rut.lower() not in ["nan", "none", ""]:
            if "-" not in rut:
                filas.append({"PROBLEMA": "RUT sin guión", "CLIENTE": row.get("Cliente"), "RUT": rut})
    det = pd.DataFrame(filas).drop_duplicates() if filas else pd.DataFrame(columns=["PROBLEMA", "CLIENTE", "RUT"])
    p(f"  RUT sin guión: {len(det)}")
    return det


def validar_destinos_inferidos(df, p):
    p("\n── V3: Destino inferido ──")
    if "Destino inferido" not in df.columns:
        return pd.DataFrame({"PROBLEMA": ["No existe columna Destino inferido"]})
    det = df[df["Destino inferido"].astype(str).str.strip() != "No"].copy()
    if not det.empty:
        out = det[["Fecha", "Cliente", "Destino", "Destino inferido"]].drop_duplicates()
        # Mensaje idéntico al original (dice "SUR" — ver Hallazgo 6 en el docstring del módulo).
        out.insert(0, "PROBLEMA", "Destino inferido no permitido en SUR")
        p(f"  ⚠ Hay {len(out)} destinos inferidos. En SUR debería ser 0.")
        return out
    p("  ✓ Sin destinos inferidos")
    return pd.DataFrame(columns=["PROBLEMA", "Fecha", "Cliente", "Destino", "Destino inferido"])


def validar_duplicados(df, p):
    p("\n── V4: Duplicados ──")
    if not all(c in df.columns for c in CLAVE_DUPLICADO):
        p("  ⚠ Faltan columnas para revisar duplicados")
        return pd.DataFrame()
    temp = df.copy()
    for c in CLAVE_DUPLICADO:
        temp[c] = temp[c].astype(str).str.strip()
    dup = temp[temp.duplicated(subset=CLAVE_DUPLICADO, keep=False)].copy()
    if not dup.empty:
        out = dup[CLAVE_DUPLICADO].sort_values(CLAVE_DUPLICADO)
        out.insert(0, "PROBLEMA", "Fila duplicada")
        p(f"  ⚠ Duplicados: {len(out)}")
        return out
    p("  ✓ Sin duplicados")
    return pd.DataFrame(columns=["PROBLEMA"] + CLAVE_DUPLICADO)


def validar_fechas(df, p):
    p("\n── V5: Fechas ──")
    temp = df.copy()
    temp["Fecha"] = pd.to_datetime(temp["Fecha"], errors="coerce")
    filas = []
    n_nulas    = int(temp["Fecha"].isna().sum())
    n_antiguas = int(((temp["Fecha"].notna()) & (temp["Fecha"].dt.year < AÑO_DESDE)).sum())
    n_futuras  = int((temp["Fecha"] > datetime.now()).sum())
    if n_nulas:
        filas.append({"PROBLEMA": "Fecha inválida", "CANTIDAD": n_nulas})
    if n_antiguas:
        filas.append({"PROBLEMA": f"Fecha anterior a {AÑO_DESDE}", "CANTIDAD": n_antiguas})
    if n_futuras:
        filas.append({"PROBLEMA": "Fecha futura", "CANTIDAD": n_futuras})
    det = pd.DataFrame(filas) if filas else pd.DataFrame(columns=["PROBLEMA", "CANTIDAD"])
    p(f"  Problemas de fecha: {len(det)}")
    return det


def conteo_por_mes_region_destino(df, p):
    p("\n── V6: Conteo por mes, región y destino ──")
    temp = df.copy()
    temp["Fecha"]          = pd.to_datetime(temp["Fecha"], errors="coerce")
    temp["Año-Mes"]        = temp["Fecha"].dt.to_period("M").astype(str)
    temp["Peso neto (kg)"] = pd.to_numeric(temp["Peso neto (kg)"], errors="coerce")
    out = (
        temp.groupby(["Año-Mes", "Región", "Destino"], dropna=False)
        .agg(Filas=("Fecha", "count"), Peso_total_kg=("Peso neto (kg)", "sum"))
        .reset_index()
        .sort_values(["Año-Mes", "Región", "Destino"])
    )
    p(f"  Combinaciones: {len(out)}")
    return out


def validar_pesos(df, p):
    p("\n── V7: Pesos ──")
    pesos = pd.to_numeric(df["Peso neto (kg)"], errors="coerce")
    filas = []
    if (pesos == 0).sum() > 0:
        filas.append({"PROBLEMA": "Peso = 0", "CANTIDAD": int((pesos == 0).sum())})
    if (pesos < 0).sum() > 0:
        filas.append({"PROBLEMA": "Peso negativo", "CANTIDAD": int((pesos < 0).sum())})
    if pesos.isna().sum() > 0:
        filas.append({"PROBLEMA": "Peso vacío o no numérico", "CANTIDAD": int(pesos.isna().sum())})
    det = pd.DataFrame(filas) if filas else pd.DataFrame(columns=["PROBLEMA", "CANTIDAD"])
    p(f"  Problemas de peso: {len(det)}")
    return det


def validar_regiones(df, p):
    p("\n── V8: Regiones válidas ──")
    if "Región" not in df.columns:
        return pd.DataFrame({"PROBLEMA": ["No existe columna Región"]})
    bad = df[~df["Región"].astype(str).isin(REGIONES_VALIDAS)][["Fecha", "Cliente", "Región"]].drop_duplicates()
    if not bad.empty:
        bad.insert(0, "PROBLEMA", "Región no válida para NORTE")
        p(f"  ⚠ Regiones inválidas: {len(bad)}")
        return bad
    p("  ✓ Regiones válidas")
    return pd.DataFrame(columns=["PROBLEMA", "Fecha", "Cliente", "Región"])


def validar_movimientos(df, p):
    p("\n── V9: Movimientos válidos ──")
    if "Movimiento" not in df.columns:
        return pd.DataFrame({"PROBLEMA": ["No existe columna Movimiento"]})
    bad = df[~df["Movimiento"].astype(str).isin(MOVIMIENTOS_VALIDOS)][["Fecha", "Cliente", "Movimiento"]].drop_duplicates()
    if not bad.empty:
        # Mensaje idéntico al original (dice "SUR" — ver Hallazgo 6 en el docstring del módulo).
        bad.insert(0, "PROBLEMA", "Movimiento no válido para SUR")
        p(f"  ⚠ Movimientos inválidos: {len(bad)}")
        return bad
    p("  ✓ Movimientos válidos")
    return pd.DataFrame(columns=["PROBLEMA", "Fecha", "Cliente", "Movimiento"])


# ══════════════════════════════════════════════════════════════════
# ESCRITURA DEL EXCEL DE SALIDA
# ══════════════════════════════════════════════════════════════════
def escribir_excel(res, ruta_salida, p):
    total = res["total_alertas"]

    resumen = pd.DataFrame({
        "Validación": [
            "V0 — Estructura columnas", "V1 — Columnas críticas vacías", "V2 — RUT sin guión",
            "V3 — Destino inferido no permitido", "V4 — Duplicados", "V5 — Fechas fuera de rango",
            "V7 — Pesos inválidos", "V8 — Regiones válidas", "V9 — Movimientos válidos", "TOTAL",
        ],
        "Alertas": [
            len(res["v0"]), len(res["v1"]), len(res["v2"]), len(res["v3"]), len(res["v4"]),
            len(res["v5"]), len(res["v7"]), len(res["v8"]), len(res["v9"]), total,
        ],
        "Estado": [
            "✓ OK" if len(res["v0"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v1"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v2"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v3"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v4"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v5"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v7"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v8"]) == 0 else "⚠ Revisar",
            "✓ OK" if len(res["v9"]) == 0 else "⚠ Revisar",
            "✓ Consolidado correcto" if total == 0 else "⚠ Hay alertas",
        ],
    })

    def hoja(writer, df, name, ok):
        out = df if not df.empty else pd.DataFrame({"resultado": [ok]})
        out.to_excel(writer, sheet_name=name, index=False)

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
        hoja(writer, res["v0"], "V0_COLUMNAS",         "Columnas OK")
        hoja(writer, res["v1"], "V1_VACIOS",           "Sin vacíos críticos")
        hoja(writer, res["v2"], "V2_RUT",              "RUT OK")
        hoja(writer, res["v3"], "V3_DESTINO_INFERIDO", "Sin destinos inferidos")
        hoja(writer, res["v4"], "V4_DUPLICADOS",       "Sin duplicados")
        hoja(writer, res["v5"], "V5_FECHAS",           "Fechas OK")
        res["v6"].to_excel(writer, sheet_name="V6_CONTEO", index=False)
        hoja(writer, res["v7"], "V7_PESOS",            "Pesos OK")
        hoja(writer, res["v8"], "V8_REGIONES",         "Regiones OK")
        hoja(writer, res["v9"], "V9_MOVIMIENTOS",      "Movimientos OK")

    # Mensaje idéntico al original (dice "SUR" — ver Hallazgo 6 en el docstring del módulo).
    p(f"\n✓ Revisión SUR generada: {ruta_salida}")


# ══════════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════
def revisar(ruta_consolidado, ruta_salida=None, hoja=HOJA_DESTINO, mostrar=True):
    """Revisa un consolidado de Norte y devuelve las alertas encontradas.

    Ver docstring de la versión RM para el detalle de parámetros y de lo que
    devuelve — el patrón es idéntico.
    """
    p = Registro(mostrar=mostrar)
    ruta = Path(ruta_consolidado)

    p("Iniciando revisión consolidado NORTE...")
    p(f"Archivo a revisar: {ruta}")

    if not ruta.exists():
        raise FileNotFoundError(f"No existe el archivo a revisar: {ruta}")

    try:
        df = pd.read_excel(ruta, sheet_name=hoja)
    except ValueError:
        disponibles = pd.ExcelFile(ruta).sheet_names
        raise ValueError(
            f"El archivo {ruta.name} no tiene la hoja '{hoja}'.\n"
            f"  Hojas disponibles: {', '.join(disponibles)}"
        )

    p(f"Filas: {len(df)} | Columnas: {len(df.columns)}")

    res = {
        "v0": validar_columnas(df, p),
        "v1": validar_vacios(df, p),
        "v2": validar_rut(df, p),
        "v3": validar_destinos_inferidos(df, p),
        "v4": validar_duplicados(df, p),
        "v5": validar_fechas(df, p),
        "v6": conteo_por_mes_region_destino(df, p),
        "v7": validar_pesos(df, p),
        "v8": validar_regiones(df, p),
        "v9": validar_movimientos(df, p),
    }
    res["total_alertas"] = sum(len(res[k]) for k in ["v0", "v1", "v2", "v3", "v4", "v5", "v7", "v8", "v9"])
    res["filas"] = len(df)
    res["columnas"] = len(df.columns)
    res["archivo"] = ruta.name

    if ruta_salida:
        escribir_excel(res, ruta_salida, p)

    res["log"] = p.texto()
    return res


# ══════════════════════════════════════════════════════════════════
# USO DESDE LA TERMINAL
# ══════════════════════════════════════════════════════════════════
def _main():
    argumentos = [a for a in sys.argv[1:] if not a.startswith("--")]

    if argumentos:
        ruta = Path(argumentos[0])
        salida = ruta.parent / "REVISION_CONSOLIDADO_NORTE.xlsx"
    else:
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
        try:
            from config import archivo_prueba, ConfiguracionFaltante
        except ImportError:
            print("\n✗ No se encontró config.py. Debe estar en la carpeta src/.\n")
            sys.exit(2)

        try:
            ruta = archivo_prueba("norte")
        except ConfiguracionFaltante as e:
            print(f"\n✗ {e}\n")
            sys.exit(2)
        except Exception as e:
            print(f"\n✗ No se pudo leer la configuración: {e}\n")
            sys.exit(2)

        salida = ruta.parent / "REVISION_CONSOLIDADO_NORTE.xlsx"

    try:
        res = revisar(ruta, ruta_salida=salida)
    except (FileNotFoundError, ValueError) as e:
        print(f"\n✗ {e}\n")
        sys.exit(2)

    sys.exit(0 if res["total_alertas"] == 0 else 1)


if __name__ == "__main__":
    _main()
