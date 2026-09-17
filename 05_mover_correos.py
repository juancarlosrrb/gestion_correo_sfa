# -*- coding: utf-8 -*-
"""
PASO 05 - Mover los correos a la carpeta de su area.

Que hace:
  - Lee predicciones_bandeja.xlsx (el que tu ya revisaste).
  - Mueve a su carpeta solo las filas que digan mover = SI.
  - Escribe el resultado de cada movimiento en la columna 'estado' y
    guarda una bitacora aparte.

Seguridad, porque esto SI toca el correo de verdad:
  - Por defecto corre en MODO SIMULACION: te muestra que haria y no mueve
    nada. Para mover de verdad hay que pasar --ejecutar.
  - Se puede exigir un umbral mas alto con --umbral.
  - Si una carpeta destino no existe, avisa y no crea nada.

Uso:
    py 05_mover_correos.py                      (simulacion, no mueve nada)
    py 05_mover_correos.py --ejecutar           (mueve de verdad)
    py 05_mover_correos.py --ejecutar --umbral 0.9
"""
import argparse
import os
import sys
import time
from datetime import datetime

sys.path.insert(0, ".")

import pandas as pd

from correo_automation.connection import connect_outlook

from modelo_correos import config
from modelo_correos.progreso import barra


def leer_argumentos():
    p = argparse.ArgumentParser(description="Mueve los correos segun el Excel de predicciones.")
    p.add_argument("--archivo", type=str, default=config.RUTA_PREDICCIONES_EXCEL,
                   help="Excel (o parquet) con las predicciones ya revisadas.")
    p.add_argument("--cuenta", type=str, default=config.EMAIL_OBJETIVO)
    p.add_argument("--ejecutar", action="store_true",
                   help="MUEVE DE VERDAD. Sin esta bandera solo simula.")
    p.add_argument("--umbral", type=float, default=0.0,
                   help="Exigir una probabilidad minima adicional (0 = usar lo que diga el Excel).")
    p.add_argument("--marcar-leido", action="store_true",
                   help="Marcar como leido el correo al moverlo.")
    p.add_argument("--pausa", type=float, default=0.4,
                   help="Segundos de pausa entre movimientos (el servidor de correo es lento).")
    return p.parse_args()


def cargar_predicciones(ruta):
    if not os.path.exists(ruta):
        print(f"ERROR: no existe {ruta}. Corre primero 04_predecir_bandeja.py")
        return None
    if ruta.lower().endswith(".parquet"):
        return pd.read_parquet(ruta)
    return pd.read_excel(ruta, sheet_name="predicciones")


def buscar_carpeta(cuenta, nombre):
    """Busca la carpeta por nombre, primero al primer nivel y despues adentro."""
    try:
        return cuenta.Folders(nombre)
    except Exception:
        pass
    # Buscar un nivel mas adentro (por si las areas estan dentro de la bandeja)
    try:
        for f in cuenta.Folders:
            try:
                for sub in f.Folders:
                    if str(sub.Name).strip().lower() == str(nombre).strip().lower():
                        return sub
            except Exception:
                continue
    except Exception:
        pass
    return None


def main():
    args = leer_argumentos()
    config.asegurar_carpetas()

    print("=" * 70)
    print("PASO 05 - MOVER CORREOS A SU CARPETA")
    print("=" * 70)

    if not args.ejecutar:
        print("\n*** MODO SIMULACION: no se va a mover nada. ***")
        print("*** Para mover de verdad agrega  --ejecutar  al final. ***\n")

    df = cargar_predicciones(args.archivo)
    if df is None or df.empty:
        return

    print(f"Filas en el archivo: {len(df)}")

    # ---------------------------------------------------------------
    # 1. Filtrar que se mueve
    # ---------------------------------------------------------------
    df["mover"] = df["mover"].astype(str).str.strip().str.upper()
    candidatos = df[df["mover"] == "SI"].copy()

    if args.umbral > 0:
        antes = len(candidatos)
        candidatos = candidatos[candidatos["probabilidad"].astype(float) >= args.umbral]
        print(f"Umbral extra {args.umbral:.2f}: de {antes} quedan {len(candidatos)}")

    candidatos = candidatos[candidatos["carpeta_destino"].astype(str).str.strip() != ""]

    if candidatos.empty:
        print("No hay correos marcados para mover. Revisa la columna 'mover' del Excel.")
        return

    print(f"\nCorreos a mover: {len(candidatos)}")
    print("\nDistribucion:")
    for carpeta, n in candidatos["carpeta_destino"].value_counts().items():
        print(f"   {n:5d}  ->  {carpeta}")

    # ---------------------------------------------------------------
    # 2. Conectar y verificar carpetas
    # ---------------------------------------------------------------
    print(f"\nConectando a Outlook (cuenta: {args.cuenta})...")
    session, cuenta = connect_outlook(args.cuenta)
    if cuenta is None:
        print("ERROR: no se encontro la cuenta.")
        return

    carpetas = {}
    faltan = []
    for nombre in candidatos["carpeta_destino"].unique():
        objeto = buscar_carpeta(cuenta, nombre)
        if objeto is None:
            faltan.append(nombre)
        else:
            carpetas[nombre] = objeto

    if faltan:
        print("\n! Estas carpetas NO existen en Outlook:")
        for n in faltan:
            print(f"     {n}")
        print("  Creala a mano en Outlook o corrige config_areas.json.")
        print("  Los correos de esas carpetas se van a saltar.")
        print("\n  Carpetas que SI existen en la cuenta:")
        try:
            for f in cuenta.Folders:
                print(f"     - {f.Name}")
        except Exception:
            pass

    candidatos = candidatos[~candidatos["carpeta_destino"].isin(faltan)]
    if candidatos.empty:
        print("\nNo queda nada que mover.")
        return

    # ---------------------------------------------------------------
    # 3. Mover
    # ---------------------------------------------------------------
    print(f"\n{'SIMULANDO' if not args.ejecutar else 'MOVIENDO'} {len(candidatos)} correos...\n")

    resultados = []
    movidos = 0
    fallidos = 0

    pbar = barra(len(candidatos), "Moviendo")
    try:
        for _, fila in candidatos.iterrows():
            entry_id = str(fila.get("entry_id", "")).strip()
            destino_nombre = str(fila["carpeta_destino"])
            asunto = str(fila.get("asunto", ""))[:60]

            if not entry_id or entry_id.lower() == "nan":
                resultados.append({**fila.to_dict(), "estado": "SIN_ENTRY_ID"})
                fallidos += 1
                pbar.update(1)
                continue

            if not args.ejecutar:
                resultados.append({**fila.to_dict(), "estado": f"SIMULADO -> {destino_nombre}"})
                pbar.update(1)
                continue

            try:
                # GetItemFromID vuelve a abrir el correo exacto por su EntryID
                m = session.GetItemFromID(entry_id)

                if args.marcar_leido:
                    try:
                        m.UnRead = False
                        m.Save()
                    except Exception:
                        pass

                m.Move(carpetas[destino_nombre])
                resultados.append({**fila.to_dict(), "estado": f"MOVIDO -> {destino_nombre}"})
                movidos += 1
                time.sleep(args.pausa)

            except Exception as e:
                resultados.append({**fila.to_dict(), "estado": f"ERROR: {str(e)[:120]}"})
                fallidos += 1
                print(f"\n   ! No se pudo mover '{asunto}': {e}")

            pbar.update(1)
    finally:
        pbar.close()

    # ---------------------------------------------------------------
    # 4. Bitacora
    # ---------------------------------------------------------------
    df_res = pd.DataFrame(resultados)
    df_res["fecha_movimiento"] = datetime.now().isoformat(timespec="seconds")

    if os.path.exists(config.RUTA_BITACORA_MOVIDOS):
        try:
            previa = pd.read_excel(config.RUTA_BITACORA_MOVIDOS)
            df_res = pd.concat([previa, df_res], ignore_index=True)
        except Exception:
            pass

    df_res.to_excel(config.RUTA_BITACORA_MOVIDOS, index=False)

    print("\n" + "=" * 70)
    print("RESUMEN")
    print("=" * 70)
    if args.ejecutar:
        print(f"Movidos correctamente: {movidos}")
        print(f"Con error            : {fallidos}")
    else:
        print(f"Se habrian movido    : {len(candidatos)}")
        print("Nada se movio (modo simulacion). Agrega --ejecutar cuando estes listo.")
    print(f"Bitacora: {config.RUTA_BITACORA_MOVIDOS}")


if __name__ == "__main__":
    main()
