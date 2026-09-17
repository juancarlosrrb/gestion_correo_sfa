# -*- coding: utf-8 -*-
"""
Guardado incremental en parquet + armado del texto que ve el modelo.

Esta es la pieza que hace que manana puedas correr el script otra vez
y no vuelva a leer los mismos correos.
"""
import os

import pandas as pd

from . import config

COLUMNAS_BASE = [
    "email_id",
    "entry_id",
    "conversation_id",
    "fecha",
    "remitente",
    "destinatario",
    "destinatarios_todos",
    "asunto",
    "cuerpo",
    "nombres_adjuntos",
    "texto_adjuntos",
    "n_adjuntos",
]


def cargar_base(ruta):
    """Lee el parquet. Si no existe todavia, devuelve un DataFrame vacio."""
    if os.path.exists(ruta):
        try:
            df = pd.read_parquet(ruta)
            print(f"Base existente cargada: {ruta} ({len(df)} correos)")
            return df
        except Exception as e:
            print(f"! No se pudo leer {ruta}: {e}")
            print("  Se creara una base nueva (el archivo viejo NO se borra).")
    else:
        print(f"La base {ruta} no existe. Se creara por primera vez.")
    return pd.DataFrame(columns=COLUMNAS_BASE)


def ids_ya_leidos(df):
    """Set con los email_id ya procesados. Set = busqueda instantanea."""
    if df is None or df.empty or "email_id" not in df.columns:
        return set()
    return set(df["email_id"].astype(str).tolist())


def guardar_base(df_existente, filas_nuevas, ruta):
    """
    Junta lo viejo con lo nuevo, quita repetidos por email_id y graba.
    Devuelve el DataFrame resultante.
    """
    if not filas_nuevas:
        return df_existente

    df_nuevo = pd.DataFrame(filas_nuevas)

    if df_existente is None or df_existente.empty:
        df_final = df_nuevo
    else:
        df_final = pd.concat([df_existente, df_nuevo], ignore_index=True)

    df_final["email_id"] = df_final["email_id"].astype(str)
    df_final = df_final.drop_duplicates(subset=["email_id"], keep="last")

    # Todo a texto: parquet se queja si una columna mezcla tipos
    for col in df_final.columns:
        if col != "n_adjuntos":
            df_final[col] = df_final[col].astype(str)

    os.makedirs(os.path.dirname(ruta) or ".", exist_ok=True)
    df_final.to_parquet(ruta, index=False)
    return df_final


def construir_texto(fila, peso_asunto=None):
    """
    Arma el texto unico que se convierte en numeros.

    OJO: esta funcion la usan el script 02 (entrenamiento) y el 04
    (prediccion). Si cambias algo aca, tienes que volver a entrenar,
    porque el modelo quedaria viendo un texto distinto al que aprendio.
    """
    if peso_asunto is None:
        peso_asunto = config.PESO_ASUNTO

    def campo(nombre):
        valor = fila.get(nombre, "")
        if valor is None or str(valor).lower() == "nan":
            return ""
        return str(valor)

    asunto = campo("asunto")

    partes = []
    # El asunto se repite: pesa mas que el resto porque es donde
    # tu hermano lee la clave para decidir a que area va.
    partes.extend([asunto] * max(1, int(peso_asunto)))
    partes.append(campo("nombres_adjuntos").replace("|", " "))
    partes.append(campo("cuerpo"))
    partes.append(campo("texto_adjuntos"))

    return " ".join(p for p in partes if p).lower()


def agregar_columna_texto(df, peso_asunto=None):
    """Agrega la columna 'texto' a un DataFrame completo."""
    df = df.copy()
    df["texto"] = df.apply(lambda f: construir_texto(f, peso_asunto), axis=1)
    return df
