# -*- coding: utf-8 -*-
"""
PASO 04 - Predecir a que area va cada correo de la Bandeja de Entrada.

Que hace:
  - Carga el vectorizador (paso 02) y el modelo (paso 03).
  - Lee los primeros N correos de la Bandeja de Entrada (N se cambia con
    --cantidad, por defecto 100).
  - Para cada uno predice el destinatario y con que probabilidad.
  - Traduce ese destinatario a una carpeta (area1, area2, ...) usando
    config_areas.json.
  - Guarda predicciones_bandeja.xlsx y .parquet.

IMPORTANTE: este script NO mueve nada. Solo propone.
Tu revisas el Excel, corriges lo que este mal, y despues corres el 05.

Uso:
    py 04_predecir_bandeja.py
    py 04_predecir_bandeja.py --cantidad 300
    py 04_predecir_bandeja.py --cantidad 50 --todos
"""
import argparse
import json
import os
import sys

sys.path.insert(0, ".")

import joblib
import numpy as np
import pandas as pd

from correo_automation.connection import connect_outlook

from modelo_correos import config
from modelo_correos.almacenamiento import construir_texto
from modelo_correos.progreso import barra
from modelo_correos.utils_outlook import extraer_correo


def leer_argumentos():
    p = argparse.ArgumentParser(description="Predice el area de los correos de la bandeja.")
    p.add_argument("--cantidad", type=int, default=100,
                   help="Cuantos correos leer de la bandeja (por defecto 100).")
    p.add_argument("--carpeta", type=str, default=config.CARPETA_ENTRADA,
                   help='Carpeta a leer (por defecto "Bandeja de Entrada").')
    p.add_argument("--cuenta", type=str, default=config.EMAIL_OBJETIVO)
    p.add_argument("--todos", action="store_true",
                   help="Leer tambien los correos ya leidos. Por defecto solo los NO leidos.")
    p.add_argument("--ocr", action="store_true", help="OCR a PDF escaneados (lento).")
    p.add_argument("--sin-adjuntos", action="store_true",
                   help="No abrir adjuntos. Mas rapido pero pierde informacion.")
    return p.parse_args()


def cargar_mapa_areas():
    """destinatario -> {area, carpeta}. Si no existe el json, lo crea vacio."""
    if not os.path.exists(config.RUTA_CONFIG_AREAS):
        print(f"! No existe {config.RUTA_CONFIG_AREAS}. Se crea una plantilla.")
        plantilla = {
            "areas": {
                "correo_del_area_1@empresa.com": {"area": "area1", "carpeta": "area1"},
                "correo_del_area_2@empresa.com": {"area": "area2", "carpeta": "area2"},
            },
            "umbral_confianza": 0.70,
            "_nota": "Edita este archivo: la llave es el correo destino, el valor dice a que carpeta de Outlook se mueve.",
        }
        with open(config.RUTA_CONFIG_AREAS, "w", encoding="utf-8") as f:
            json.dump(plantilla, f, ensure_ascii=False, indent=2)
        return plantilla

    with open(config.RUTA_CONFIG_AREAS, "r", encoding="utf-8") as f:
        return json.load(f)


def main():
    args = leer_argumentos()
    config.asegurar_carpetas()

    print("=" * 70)
    print("PASO 04 - PREDICCION SOBRE LA BANDEJA DE ENTRADA")
    print("=" * 70)

    # ---------------------------------------------------------------
    # 1. Cargar modelo y vectorizador
    # ---------------------------------------------------------------
    for ruta in (config.RUTA_VECTORIZADOR, config.RUTA_MODELO):
        if not os.path.exists(ruta):
            print(f"ERROR: falta {ruta}. Corre los pasos 02 y 03 primero.")
            return

    vectorizador = joblib.load(config.RUTA_VECTORIZADOR)
    paquete = joblib.load(config.RUTA_MODELO)
    modelo = paquete["modelo"]

    print(f"Modelo: {paquete['tipo']}, entrenado el {paquete['fecha_entrenamiento']}")
    print(f"Acierto medido en test: {paquete['acierto_test']*100:.2f}%")
    print(f"Clases que conoce: {len(paquete['clases'])}")

    cfg_areas = cargar_mapa_areas()
    mapa = cfg_areas.get("areas", {})
    umbral = float(cfg_areas.get("umbral_confianza", 0.70))
    print(f"Umbral de confianza configurado: {umbral:.2f}")

    faltantes = [c for c in paquete["clases"] if c not in mapa]
    if faltantes:
        print(f"\n! Estos destinatarios que el modelo conoce NO estan en {config.RUTA_CONFIG_AREAS}:")
        for c in faltantes:
            print(f"     {c}")
        print("  Sus correos saldran con area vacia y el paso 05 los dejara quietos.")

    # ---------------------------------------------------------------
    # 2. Leer la bandeja
    # ---------------------------------------------------------------
    print(f"\nConectando a Outlook (cuenta: {args.cuenta})...")
    session, cuenta = connect_outlook(args.cuenta)
    if cuenta is None:
        print("ERROR: no se encontro la cuenta.")
        return

    try:
        carpeta = cuenta.Folders(args.carpeta)
    except Exception as e:
        print(f"ERROR: no se pudo abrir '{args.carpeta}': {e}")
        return

    items = carpeta.Items
    items.Sort("[ReceivedTime]", True)  # del mas nuevo al mas viejo

    if not args.todos:
        try:
            items = items.Restrict("[UnRead] = True")
            items.Sort("[ReceivedTime]", True)
        except Exception:
            print("! No se pudo filtrar por no leidos, se leen todos.")

    total = items.Count
    objetivo = min(args.cantidad, total)
    print(f"Carpeta '{args.carpeta}': {total} correos disponibles. Se leeran {objetivo}.")

    filas = []
    pbar = barra(objetivo, "Leyendo bandeja")
    try:
        for i in range(1, total + 1):
            if len(filas) >= objetivo:
                break
            try:
                m = items.Item(i)
                if str(m.Class) != "43":
                    continue
                fila = extraer_correo(
                    m, usar_ocr=args.ocr, recortar=True,
                    leer_adjuntos=not args.sin_adjuntos,
                )
                filas.append(fila)
                pbar.update(1)
            except Exception as e:
                print(f"\n   ! Error en correo #{i}: {e}")
    finally:
        pbar.close()

    if not filas:
        print("No se leyo ningun correo. Nada que predecir.")
        return

    df = pd.DataFrame(filas)
    print(f"\nCorreos leidos: {len(df)}")

    # ---------------------------------------------------------------
    # 3. Predecir
    # ---------------------------------------------------------------
    print("Convirtiendo a numeros y prediciendo...")
    textos = df.apply(lambda f: construir_texto(f, paquete.get("peso_asunto")), axis=1)
    X = vectorizador.transform(textos)

    predicciones = modelo.predict(X)

    if hasattr(modelo, "predict_proba"):
        probas = modelo.predict_proba(X)
        confianza = probas.max(axis=1)
        clases = np.array(modelo.classes_)
        # Segunda opcion: util cuando la primera va justa
        orden = np.argsort(probas, axis=1)
        segunda = clases[orden[:, -2]] if probas.shape[1] > 1 else predicciones
        conf_segunda = probas[np.arange(len(probas)), orden[:, -2]] if probas.shape[1] > 1 else 0
    else:
        confianza = np.ones(len(df))
        segunda = predicciones
        conf_segunda = np.zeros(len(df))

    df_out = pd.DataFrame({
        "email_id": df["email_id"],
        "entry_id": df["entry_id"],
        "fecha": df["fecha"],
        "remitente": df["remitente"],
        "asunto": df["asunto"],
        "nombres_adjuntos": df["nombres_adjuntos"],
        "correo_escogido": predicciones,
        "probabilidad": np.round(confianza, 4),
        "segunda_opcion": segunda,
        "prob_segunda": np.round(conf_segunda, 4),
    })

    df_out["area"] = df_out["correo_escogido"].map(
        lambda c: mapa.get(c, {}).get("area", "")
    )
    df_out["carpeta_destino"] = df_out["correo_escogido"].map(
        lambda c: mapa.get(c, {}).get("carpeta", "")
    )
    df_out["mover"] = np.where(
        (df_out["probabilidad"] >= umbral) & (df_out["carpeta_destino"] != ""),
        "SI", "NO",
    )
    df_out["revisado_por_mi"] = ""   # columna para que tu hermano marque a mano
    df_out["estado"] = ""            # la llena el paso 05

    df_out = df_out.sort_values("probabilidad", ascending=False)

    # ---------------------------------------------------------------
    # 4. Guardar
    # ---------------------------------------------------------------
    df_out.to_parquet(config.RUTA_PREDICCIONES_PARQUET, index=False)
    with pd.ExcelWriter(config.RUTA_PREDICCIONES_EXCEL, engine="openpyxl") as w:
        df_out.to_excel(w, sheet_name="predicciones", index=False)
        resumen = (
            df_out.groupby(["area", "correo_escogido", "mover"])
            .size().reset_index(name="cantidad")
            .sort_values("cantidad", ascending=False)
        )
        resumen.to_excel(w, sheet_name="resumen", index=False)

    print("\n" + "=" * 70)
    print("RESUMEN DE LA PREDICCION")
    print("=" * 70)
    print(f"Correos procesados: {len(df_out)}")
    print(f"Se moverian automatico (prob >= {umbral:.2f}): "
          f"{(df_out['mover']=='SI').sum()}")
    print(f"Quedan para revisar a mano: {(df_out['mover']=='NO').sum()}")
    print("\nPor area:")
    for area, n in df_out["area"].replace("", "(sin mapear)").value_counts().items():
        print(f"   {n:5d}  {area}")

    print(f"\nExcel   : {config.RUTA_PREDICCIONES_EXCEL}")
    print(f"Parquet : {config.RUTA_PREDICCIONES_PARQUET}")
    print("\nAbre el Excel, revisa la columna 'mover' y corrige lo que veas mal.")
    print("Cuando quedes conforme:  py 05_mover_correos.py")


if __name__ == "__main__":
    main()
