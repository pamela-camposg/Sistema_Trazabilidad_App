# =============================================================================
# procesar.py — ESTE es el único archivo que conectas con tus scripts.
#
# La app le pasa a la función procesar():
#   - zona:            "RM", "SUR" o "NORTE"
#   - carpeta_entrada: donde quedaron los Excel que subió la operadora
#                      (ej. "/work/uploads/RECYNOR.xlsx")
#   - carpeta_salida:  donde debes dejar el consolidado generado
#
# Debe devolver un diccionario con:
#   resumen  -> texto que se muestra arriba
#   fuentes  -> lista de {archivo, hojas, filas} para la tabla
#   alertas  -> lista de {titulo, detalle} del control de calidad
#   salida   -> nombre del archivo consolidado que dejaste en carpeta_salida
#   log      -> texto largo con el detalle técnico
#
# Ahora mismo trae una VERSIÓN DE DEMOSTRACIÓN que sí funciona: lee cada Excel,
# reporta hojas y filas, y arma un consolidado de prueba. Sirve para ver la app
# completa de punta a punta. Cuando conectes tu lógica real de RM, reemplazas
# SOLO el bloque marcado "AQUÍ VA TU LÓGICA REAL".
# =============================================================================

import os
import pandas as pd


def _listar_archivos(carpeta):
    return sorted(
        os.path.join(carpeta, f)
        for f in os.listdir(carpeta)
        if f.lower().endswith(".xlsx")
    )


def procesar(zona, carpeta_entrada, carpeta_salida):
    rutas = _listar_archivos(carpeta_entrada)
    log = [f"Zona: {zona}", f"Archivos recibidos: {len(rutas)}", ""]
    fuentes = []
    alertas = []

    # -------------------------------------------------------------------------
    # DEMOSTRACIÓN: leer cada archivo y reportar qué trae.
    # Esto prueba que pandas + openpyxl funcionan en el navegador.
    # -------------------------------------------------------------------------
    marcos = []
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

    # =========================================================================
    # AQUÍ VA TU LÓGICA REAL (cuando conectemos RM)
    # -------------------------------------------------------------------------
    # 1. Sube tus scripts refactorizados al repo, dentro de python/rm/
    #    (consolidar.py, control_calidad.py, revisar_consolidado.py).
    # 2. Impórtalos arriba, por ejemplo:
    #         from rm import consolidar, control_calidad
    # 3. Llama a tus funciones pasándoles las rutas de carpeta_entrada.
    #    Como refactorizaste todo para recibir rutas como parámetros, tus
    #    funciones corren aquí sin cambios: solo cambian las rutas.
    #         df_final = consolidar.consolidar_rm(
    #             ruta_recynor = carpeta_entrada + "/RECYNOR.xlsx",
    #             ruta_homologacion = carpeta_entrada + "/HOMOLOGACION.xlsx",
    #             ... )
    #         qc = control_calidad.correr(...)
    # 4. Vuelca df_final a Excel en carpeta_salida y traduce el control de
    #    calidad a la lista "alertas".
    # =========================================================================

    # Consolidado de DEMOSTRACIÓN (reemplazar por df_final real)
    salida = None
    if marcos:
        df_todo = pd.concat(marcos, ignore_index=True)
        salida = f"CONSOLIDADO_{zona}.xlsx"
        df_todo.to_excel(os.path.join(carpeta_salida, salida), index=False)
        log.append("")
        log.append(f"Consolidado de demostración: {len(df_todo)} filas")
        resumen = (
            f"<b>Zona {zona}</b> · {len(rutas)} archivo(s) leídos · "
            f"{len(df_todo)} filas combinadas (versión de demostración)."
        )
    else:
        resumen = "No se leyó ningún archivo válido."

    if not alertas:
        alertas.append({
            "titulo": "Versión de demostración",
            "detalle": "Todavía no está conectada la lógica real de la zona. "
                       "El consolidado es solo una combinación de prueba.",
        })

    return {
        "resumen": resumen,
        "fuentes": fuentes,
        "alertas": alertas,
        "salida": salida,
        "log": "\n".join(log),
    }
