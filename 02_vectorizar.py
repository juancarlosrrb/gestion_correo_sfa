# -*- coding: utf-8 -*-
"""
PASO 02 - Convertir el texto de los correos en numeros.

Que hace:
  - Lee base_de_datos_1historica_modelo.parquet.
  - Arma un solo campo de texto por correo (asunto x3 + adjuntos + cuerpo).
  - Lo convierte en una matriz de numeros con TF-IDF.
  - Guarda: el vectorizador (para reusarlo al predecir), la matriz X
    y las etiquetas Y (el destinatario de cada correo).

Por que TF-IDF y no algo mas moderno:
  Con 2.000 correos al dia, texto juridico repetitivo y categorias fijas,
  TF-IDF + regresion logistica suele dar 90%+ y corre en segundos en un
  PC normal. Si mas adelante se queda corto, se cambia solo este paso
  por embeddings sin tocar los demas scripts.

Uso:
    py 02_vectorizar.py
    py 02_vectorizar.py --min-por-clase 10
"""
import argparse
import os
import sys

sys.path.insert(0, ".")

import numpy as np
import pandas as pd
from scipy import sparse
import joblib
from sklearn.feature_extraction.text import TfidfVectorizer

from modelo_correos import config
from modelo_correos.almacenamiento import agregar_columna_texto, cargar_base


def leer_argumentos():
    p = argparse.ArgumentParser(description="Convierte el texto de los correos en numeros.")
    p.add_argument("--min-por-clase", type=int, default=config.MIN_CORREOS_POR_CLASE,
                   help="Destinatarios con menos correos que esto se descartan.")
    p.add_argument("--max-features", type=int, default=30000,
                   help="Cuantas palabras/frases distintas se guardan como columnas.")
    return p.parse_args()


def main():
    args = leer_argumentos()
    config.asegurar_carpetas()

    print("=" * 70)
    print("PASO 02 - TEXTO A NUMEROS (TF-IDF)")
    print("=" * 70)

    # ---------------------------------------------------------------
    # 1. Cargar la base
    # ---------------------------------------------------------------
    if not os.path.exists(config.BASE_MODELO):
        print(f"ERROR: no existe {config.BASE_MODELO}. Corre primero 01_extraer_enviados.py")
        return

    df = cargar_base(config.BASE_MODELO)
    if df.empty:
        print("ERROR: la base esta vacia.")
        return

    print(f"Correos en la base: {len(df)}")

    # ---------------------------------------------------------------
    # 2. Limpiar clases con muy pocos correos
    # ---------------------------------------------------------------
    conteo = df["destinatario"].value_counts()
    clases_validas = conteo[conteo >= args.min_por_clase].index.tolist()
    descartadas = conteo[conteo < args.min_por_clase]

    if len(descartadas) > 0:
        print(f"\nSe descartan {len(descartadas)} destinatarios con menos de "
              f"{args.min_por_clase} correos ({descartadas.sum()} correos):")
        for correo, n in descartadas.head(10).items():
            print(f"   {n:4d}  {correo}")
        if len(descartadas) > 10:
            print(f"   ... y {len(descartadas)-10} mas")

    df = df[df["destinatario"].isin(clases_validas)].copy()

    if df.empty:
        print("ERROR: no quedo ningun correo. Baja --min-por-clase o extrae mas correos.")
        return

    print(f"\nCorreos utiles para entrenar: {len(df)}")
    print(f"Destinatarios (clases): {len(clases_validas)}")
    for correo, n in df["destinatario"].value_counts().items():
        print(f"   {n:6d}  {correo}")

    # ---------------------------------------------------------------
    # 3. Armar el texto unico por correo
    # ---------------------------------------------------------------
    print("\nArmando el campo de texto (asunto x{} + adjuntos + cuerpo)...".format(config.PESO_ASUNTO))
    df = agregar_columna_texto(df)

    vacios = (df["texto"].str.strip() == "").sum()
    if vacios:
        print(f"   ! {vacios} correos quedaron con texto vacio, se descartan.")
        df = df[df["texto"].str.strip() != ""].copy()

    # ---------------------------------------------------------------
    # 4. TF-IDF
    # ---------------------------------------------------------------
    print("\nEntrenando el vectorizador TF-IDF...")
    vectorizador = TfidfVectorizer(
        max_features=args.max_features,
        ngram_range=(1, 2),        # palabras sueltas y parejas de palabras
        min_df=2,                  # una palabra debe salir en al menos 2 correos
        max_df=0.9,                # si sale en el 90% de los correos, no distingue nada
        sublinear_tf=True,         # suaviza las palabras que se repiten mucho
        strip_accents="unicode",   # "acción" y "accion" cuentan igual
        lowercase=True,
        stop_words=config.STOP_WORDS_ES,
    )

    X = vectorizador.fit_transform(df["texto"])
    y = df["destinatario"].values

    print(f"Matriz resultante: {X.shape[0]} correos x {X.shape[1]} columnas (palabras/frases)")
    densidad = X.nnz / (X.shape[0] * X.shape[1]) * 100
    print(f"Densidad: {densidad:.3f}% (casi todo son ceros, por eso se guarda comprimida)")

    # ---------------------------------------------------------------
    # 5. Guardar
    # ---------------------------------------------------------------
    joblib.dump(vectorizador, config.RUTA_VECTORIZADOR)
    sparse.save_npz(config.RUTA_MATRIZ_X, X)

    df_y = pd.DataFrame({
        "email_id": df["email_id"].values,
        "destinatario": y,
        "asunto": df["asunto"].values,
    })
    df_y.to_parquet(config.RUTA_ETIQUETAS_Y, index=False)

    print("\nGuardado:")
    print(f"   Vectorizador : {config.RUTA_VECTORIZADOR}")
    print(f"   Matriz X     : {config.RUTA_MATRIZ_X}")
    print(f"   Etiquetas Y  : {config.RUTA_ETIQUETAS_Y}")

    # ---------------------------------------------------------------
    # 6. Curiosidad util: las palabras mas caracteristicas de cada area
    # ---------------------------------------------------------------
    print("\nPalabras con mas peso en cada destinatario (promedio TF-IDF):")
    vocabulario = np.array(vectorizador.get_feature_names_out())
    for destinatario in pd.Series(y).value_counts().index[:6]:
        mascara = (y == destinatario)
        promedios = np.asarray(X[mascara].mean(axis=0)).ravel()
        top = promedios.argsort()[-8:][::-1]
        palabras = ", ".join(vocabulario[top])
        print(f"   {destinatario}")
        print(f"      -> {palabras}")

    print("\nListo. Sigue:  py 03_entrenar_modelo.py")


if __name__ == "__main__":
    main()
