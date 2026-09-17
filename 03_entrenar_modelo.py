# -*- coding: utf-8 -*-
"""
PASO 03 - Entrenar el modelo y medir que tan bien predice.

Que hace:
  - Lee la matriz X y las etiquetas Y del paso 02.
  - Parte los datos: 80% para entrenar, 20% para EVALUAR.
    El 20% de test el modelo NO lo ve al entrenar. Por eso sirve para
    saber si de verdad aprendio o solo se memorizo los correos.
  - Entrena una regresion logistica (da probabilidades, que es lo que
    necesitas para decidir cuales mover solo y cuales revisar a mano).
  - Guarda el modelo y un Excel con los resultados, incluido el detalle
    correo por correo de los que se equivoco.

Uso:
    py 03_entrenar_modelo.py
    py 03_entrenar_modelo.py --test 0.3
    py 03_entrenar_modelo.py --modelo svm
"""
import argparse
import json
import os
import sys
from datetime import datetime

sys.path.insert(0, ".")

import numpy as np
import pandas as pd
from scipy import sparse
import joblib
from sklearn.calibration import CalibratedClassifierCV
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score,
    classification_report,
    confusion_matrix,
)
from sklearn.model_selection import cross_val_score, train_test_split
from sklearn.naive_bayes import ComplementNB
from sklearn.svm import LinearSVC

from modelo_correos import config


def leer_argumentos():
    p = argparse.ArgumentParser(description="Entrena el clasificador de correos.")
    p.add_argument("--test", type=float, default=config.PROPORCION_TEST,
                   help="Proporcion de correos reservados para evaluar (0.2 = 20%%).")
    p.add_argument("--modelo", type=str, default="logistica",
                   choices=["logistica", "svm", "bayes"],
                   help="Algoritmo a usar.")
    p.add_argument("--validacion-cruzada", action="store_true",
                   help="Ademas del test, hace validacion cruzada de 5 pliegues (mas lento, mas confiable).")
    return p.parse_args()


def construir_modelo(nombre):
    if nombre == "logistica":
        # Da probabilidades directo. Es el que mejor te sirve.
        return LogisticRegression(
            max_iter=1000, class_weight="balanced", C=5.0,
        )
    if nombre == "svm":
        # Suele ganar en texto, pero no da probabilidades:
        # lo envuelvo en un calibrador para que si las de.
        return CalibratedClassifierCV(
            LinearSVC(class_weight="balanced", C=1.0), cv=3,
        )
    return ComplementNB(alpha=0.3)  # bayes: rapidisimo, buena linea base


def main():
    args = leer_argumentos()
    config.asegurar_carpetas()

    print("=" * 70)
    print("PASO 03 - ENTRENAMIENTO DEL MODELO")
    print("=" * 70)

    if not os.path.exists(config.RUTA_MATRIZ_X):
        print(f"ERROR: falta {config.RUTA_MATRIZ_X}. Corre primero 02_vectorizar.py")
        return

    X = sparse.load_npz(config.RUTA_MATRIZ_X)
    df_y = pd.read_parquet(config.RUTA_ETIQUETAS_Y)
    y = df_y["destinatario"].values

    print(f"Correos          : {X.shape[0]}")
    print(f"Columnas (rasgos): {X.shape[1]}")
    print(f"Clases           : {len(np.unique(y))}")

    # ---------------------------------------------------------------
    # 1. Partir en entrenamiento y test
    # ---------------------------------------------------------------
    conteo = pd.Series(y).value_counts()
    estratificar = y if conteo.min() >= 2 else None
    if estratificar is None:
        print("\n! Hay clases con un solo correo: no se puede estratificar el corte.")

    indices = np.arange(X.shape[0])
    idx_tr, idx_te = train_test_split(
        indices, test_size=args.test, random_state=config.SEMILLA, stratify=estratificar,
    )

    X_tr, X_te = X[idx_tr], X[idx_te]
    y_tr, y_te = y[idx_tr], y[idx_te]

    print(f"\nEntrenamiento: {len(y_tr)} correos")
    print(f"Test (no los ve al entrenar): {len(y_te)} correos")

    # ---------------------------------------------------------------
    # 2. Entrenar
    # ---------------------------------------------------------------
    print(f"\nEntrenando modelo '{args.modelo}'...")
    modelo = construir_modelo(args.modelo)
    modelo.fit(X_tr, y_tr)
    print("Entrenado.")

    # ---------------------------------------------------------------
    # 3. Evaluar
    # ---------------------------------------------------------------
    pred_tr = modelo.predict(X_tr)
    pred_te = modelo.predict(X_te)

    acc_tr = accuracy_score(y_tr, pred_tr)
    acc_te = accuracy_score(y_te, pred_te)

    print("\n" + "=" * 70)
    print("RESULTADOS")
    print("=" * 70)
    print(f"Acierto en entrenamiento : {acc_tr*100:.2f}%")
    print(f"Acierto en TEST          : {acc_te*100:.2f}%   <-- este es el que importa")

    if acc_tr - acc_te > 0.15:
        print("\n! OJO: el modelo acierta mucho mas en entrenamiento que en test.")
        print("  Se esta memorizando (sobreajuste). Necesitas mas correos por clase.")

    probas_te = None
    if hasattr(modelo, "predict_proba"):
        probas_te = modelo.predict_proba(X_te)
        conf_max = probas_te.max(axis=1)
        aciertos = (pred_te == y_te)
        print(f"\nConfianza promedio cuando ACIERTA    : {conf_max[aciertos].mean()*100:.1f}%")
        if (~aciertos).sum() > 0:
            print(f"Confianza promedio cuando SE EQUIVOCA: {conf_max[~aciertos].mean()*100:.1f}%")

        print("\nQue pasa si solo mueves automatico los de alta confianza:")
        print("   umbral    cuantos mueve    acierto de esos    quedan a mano")
        for umbral in (0.50, 0.60, 0.70, 0.80, 0.90, 0.95):
            mascara = conf_max >= umbral
            if mascara.sum() == 0:
                continue
            precision = aciertos[mascara].mean()
            print(f"   {umbral:.2f}      {mascara.sum():5d} ({mascara.mean()*100:4.1f}%)"
                  f"      {precision*100:6.2f}%          {(~mascara).sum():5d}")

    print("\nDetalle por destinatario:")
    print(classification_report(y_te, pred_te, zero_division=0))

    if args.validacion_cruzada:
        print("Validacion cruzada de 5 pliegues (paciencia)...")
        puntajes = cross_val_score(construir_modelo(args.modelo), X, y, cv=5, n_jobs=-1)
        print(f"   Acierto: {puntajes.mean()*100:.2f}% (+/- {puntajes.std()*100:.2f}%)")

    # ---------------------------------------------------------------
    # 4. Guardar modelo
    # ---------------------------------------------------------------
    paquete = {
        "modelo": modelo,
        "clases": list(modelo.classes_),
        "tipo": args.modelo,
        "acierto_test": float(acc_te),
        "acierto_train": float(acc_tr),
        "n_entrenamiento": int(len(y_tr)),
        "n_test": int(len(y_te)),
        "fecha_entrenamiento": datetime.now().isoformat(timespec="seconds"),
        "peso_asunto": config.PESO_ASUNTO,
    }
    joblib.dump(paquete, config.RUTA_MODELO)
    print(f"\nModelo guardado: {config.RUTA_MODELO}")

    # ---------------------------------------------------------------
    # 5. Guardar resultados en Excel
    # ---------------------------------------------------------------
    df_test = pd.DataFrame({
        "email_id": df_y["email_id"].values[idx_te],
        "asunto": df_y["asunto"].values[idx_te],
        "destinatario_real": y_te,
        "destinatario_predicho": pred_te,
        "acerto": (pred_te == y_te),
    })
    if probas_te is not None:
        df_test["probabilidad"] = probas_te.max(axis=1).round(4)
        df_test = df_test.sort_values("probabilidad", ascending=False)

    df_resumen = pd.DataFrame([
        {"metrica": "acierto_test", "valor": round(acc_te, 4)},
        {"metrica": "acierto_entrenamiento", "valor": round(acc_tr, 4)},
        {"metrica": "correos_entrenamiento", "valor": len(y_tr)},
        {"metrica": "correos_test", "valor": len(y_te)},
        {"metrica": "clases", "valor": len(np.unique(y))},
        {"metrica": "tipo_modelo", "valor": args.modelo},
        {"metrica": "fecha", "valor": paquete["fecha_entrenamiento"]},
    ])

    reporte = classification_report(y_te, pred_te, zero_division=0, output_dict=True)
    df_reporte = pd.DataFrame(reporte).T.round(4)

    etiquetas = sorted(np.unique(np.concatenate([y_te, pred_te])))
    df_confusion = pd.DataFrame(
        confusion_matrix(y_te, pred_te, labels=etiquetas),
        index=[f"real: {e}" for e in etiquetas],
        columns=[f"pred: {e}" for e in etiquetas],
    )

    with pd.ExcelWriter(config.RUTA_RESULTADOS_EXCEL, engine="openpyxl") as w:
        df_resumen.to_excel(w, sheet_name="1_resumen", index=False)
        df_reporte.to_excel(w, sheet_name="2_por_destinatario")
        df_confusion.to_excel(w, sheet_name="3_matriz_confusion")
        df_test.to_excel(w, sheet_name="4_detalle_test", index=False)
        df_test[~df_test["acerto"]].to_excel(w, sheet_name="5_errores", index=False)

    with open(config.RUTA_RESULTADOS_JSON, "w", encoding="utf-8") as f:
        json.dump(
            {k: v for k, v in paquete.items() if k not in ("modelo",)},
            f, ensure_ascii=False, indent=2,
        )

    print(f"Resultados: {config.RUTA_RESULTADOS_EXCEL}")
    print("   Revisa la hoja '5_errores': ahi ves en que se equivoco y por que.")
    print("   La hoja '3_matriz_confusion' te dice que areas se confunden entre si.")
    print("\nSigue:  py 04_predecir_bandeja.py")


if __name__ == "__main__":
    main()
