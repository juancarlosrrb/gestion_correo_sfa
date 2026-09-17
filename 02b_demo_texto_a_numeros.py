# -*- coding: utf-8 -*-
"""
PASO 02b - DEMO PARA APRENDER (no es parte del pipeline).

Toma los 10 correos de base_de_datos_0historica_cruda_modelo.parquet y
te muestra, paso a paso y en un Excel, como un texto se vuelve numeros.

Genera demo_texto_a_numeros.xlsx con estas hojas:
  1_correos_crudos   : los 10 correos como salieron de Outlook
  2_texto_armado     : el texto unico que ve el modelo
  3_conteo           : cuantas veces sale cada palabra (CountVectorizer)
  4_tfidf            : el peso de cada palabra (TfidfVectorizer)
  5_explicacion_idf  : por que unas palabras pesan mas que otras

Si la base cruda todavia no existe, corre con 10 correos de ejemplo
inventados, para que igual puedas ver el mecanismo.

Uso:
    py 02b_demo_texto_a_numeros.py
"""
import os
import sys

sys.path.insert(0, ".")

import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer

from modelo_correos import config
from modelo_correos.almacenamiento import agregar_columna_texto, cargar_base

EJEMPLOS = [
    ("area1@empresa.com", "Accion de tutela FOMAG docente", "Se notifica accion de tutela contra el fondo de prestaciones del magisterio"),
    ("area1@empresa.com", "Tutela FOMAG incidente desacato", "Incidente de desacato dentro de la accion de tutela del magisterio"),
    ("area1@empresa.com", "Notificacion tutela magisterio", "Juzgado notifica tutela FOMAG solicitud de respuesta urgente"),
    ("area2@empresa.com", "Embargo cuenta bancaria proceso ejecutivo", "Se ordena medida cautelar de embargo sobre cuenta del demandado"),
    ("area2@empresa.com", "Proceso ejecutivo mandamiento de pago", "El juzgado libra mandamiento de pago dentro del proceso ejecutivo"),
    ("area2@empresa.com", "Medida cautelar embargo de salarios", "Oficio de embargo de salarios dentro del ejecutivo singular"),
    ("area3@empresa.com", "Derecho de peticion respuesta", "Se radica derecho de peticion solicitando informacion del afiliado"),
    ("area3@empresa.com", "Peticion informacion afiliado", "El usuario radica peticion de informacion sobre su historia laboral"),
    ("area3@empresa.com", "Respuesta derecho de peticion radicado", "Respuesta al derecho de peticion radicado por el ciudadano"),
    ("area1@empresa.com", "Tutela FOMAG fallo primera instancia", "Fallo de primera instancia en accion de tutela contra el FOMAG"),
]


def cargar_datos_demo():
    if os.path.exists(config.BASE_CRUDA):
        df = cargar_base(config.BASE_CRUDA)
        if not df.empty:
            print(f"Usando la base cruda real: {len(df)} correos")
            return df.head(10).reset_index(drop=True), True

    print("No existe la base cruda todavia -> uso 10 correos de ejemplo inventados.")
    print("(Corre 01_extraer_enviados.py para verlo con tus correos reales.)\n")
    df = pd.DataFrame([
        {
            "email_id": f"demo-{i+1}",
            "entry_id": "",
            "conversation_id": "",
            "fecha": "",
            "remitente": "gestor@empresa.com",
            "destinatario": dest,
            "destinatarios_todos": dest,
            "asunto": asunto,
            "cuerpo": cuerpo,
            "nombres_adjuntos": "",
            "texto_adjuntos": "",
            "n_adjuntos": 0,
        }
        for i, (dest, asunto, cuerpo) in enumerate(EJEMPLOS)
    ])
    return df, False


def main():
    config.asegurar_carpetas()

    print("=" * 70)
    print("DEMO - COMO SE CONVIERTE UN CORREO EN NUMEROS")
    print("=" * 70)

    df, es_real = cargar_datos_demo()

    # -----------------------------------------------------------------
    # PASO A: de varios campos a un solo texto
    # -----------------------------------------------------------------
    print("\n--- PASO A: juntar asunto + adjuntos + cuerpo en un solo texto ---")
    print(f"El asunto se repite {config.PESO_ASUNTO} veces a proposito: asi pesa mas.\n")

    df = agregar_columna_texto(df)
    print("Ejemplo (correo 1):")
    print(f"   asunto : {str(df.loc[0,'asunto'])[:70]}")
    print(f"   texto  : {str(df.loc[0,'texto'])[:150]}...")

    textos = df["texto"].tolist()
    etiquetas = [f"correo_{i+1}" for i in range(len(df))]

    # -----------------------------------------------------------------
    # PASO B: CountVectorizer, contar palabras
    # -----------------------------------------------------------------
    print("\n--- PASO B: contar palabras (CountVectorizer) ---")
    print("Cada palabra distinta se vuelve una COLUMNA.")
    print("Cada correo es una FILA. El numero = cuantas veces sale esa palabra.\n")

    contador = CountVectorizer(
        strip_accents="unicode", lowercase=True,
        stop_words=config.STOP_WORDS_ES, min_df=1,
    )
    X_conteo = contador.fit_transform(textos)
    vocab = contador.get_feature_names_out()

    df_conteo = pd.DataFrame(X_conteo.toarray(), columns=vocab, index=etiquetas)
    print(f"Vocabulario: {len(vocab)} palabras distintas")
    print(f"Matriz: {df_conteo.shape[0]} filas (correos) x {df_conteo.shape[1]} columnas (palabras)")

    columnas_muestra = list(vocab[:8])
    print(f"\nPrimeras columnas ({', '.join(columnas_muestra)}):")
    print(df_conteo[columnas_muestra].head(5).to_string())

    # -----------------------------------------------------------------
    # PASO C: TF-IDF, pesar palabras
    # -----------------------------------------------------------------
    print("\n--- PASO C: pesar palabras (TF-IDF) ---")
    print("Problema del conteo: 'juzgado' sale en TODOS los correos, no distingue nada.")
    print("TF-IDF castiga las palabras comunes y premia las raras.")
    print("   TF  = que tan seguido sale la palabra en ESTE correo")
    print("   IDF = que tan rara es la palabra en TODOS los correos")
    print("   peso = TF x IDF\n")

    tfidf = TfidfVectorizer(
        strip_accents="unicode", lowercase=True,
        stop_words=config.STOP_WORDS_ES, min_df=1,
    )
    X_tfidf = tfidf.fit_transform(textos)
    vocab_tfidf = tfidf.get_feature_names_out()
    df_tfidf = pd.DataFrame(
        np.round(X_tfidf.toarray(), 4), columns=vocab_tfidf, index=etiquetas
    )

    print("Las mismas columnas, ahora con peso en vez de conteo:")
    print(df_tfidf[[c for c in columnas_muestra if c in df_tfidf.columns]].head(5).to_string())

    # -----------------------------------------------------------------
    # PASO D: el IDF, palabra por palabra
    # -----------------------------------------------------------------
    print("\n--- PASO D: que palabras pesan mas y por que ---")
    df_idf = pd.DataFrame({
        "palabra": vocab_tfidf,
        "idf": np.round(tfidf.idf_, 4),
        "en_cuantos_correos_sale": np.asarray((X_conteo > 0).sum(axis=0)).ravel()
        if len(vocab) == len(vocab_tfidf) else 0,
    }).sort_values("idf", ascending=False)

    print("\nLas 8 palabras que MAS pesan (raras, muy informativas):")
    print(df_idf.head(8).to_string(index=False))
    print("\nLas 8 que MENOS pesan (salen en casi todos, poco utiles):")
    print(df_idf.tail(8).to_string(index=False))

    # -----------------------------------------------------------------
    # PASO E: a que se parece cada correo
    # -----------------------------------------------------------------
    print("\n--- PASO E: para que sirve todo esto ---")
    print("Ya convertido en numeros, el computador puede medir que tan parecidos")
    print("son dos correos. Si el correo nuevo se parece a los que fueron al area1,")
    print("el modelo predice area1. Eso es todo el truco.\n")

    from sklearn.metrics.pairwise import cosine_similarity
    similitud = pd.DataFrame(
        np.round(cosine_similarity(X_tfidf), 3), index=etiquetas, columns=etiquetas
    )
    print("Parecido entre los primeros 5 correos (1.0 = identicos, 0.0 = nada que ver):")
    print(similitud.iloc[:5, :5].to_string())

    print("\nDestinatario real de cada uno:")
    for i, dest in enumerate(df["destinatario"].head(5)):
        print(f"   correo_{i+1}: {dest}")
    print("\nFijate: los correos del mismo destinatario tienen parecido alto entre si.")

    # -----------------------------------------------------------------
    # Guardar el Excel
    # -----------------------------------------------------------------
    cols_crudas = [c for c in ["email_id", "destinatario", "asunto", "cuerpo",
                               "nombres_adjuntos", "texto_adjuntos", "n_adjuntos"]
                   if c in df.columns]

    with pd.ExcelWriter(config.RUTA_DEMO_EXCEL, engine="openpyxl") as w:
        df[cols_crudas].to_excel(w, sheet_name="1_correos_crudos", index=False)
        df[["destinatario", "asunto", "texto"]].to_excel(w, sheet_name="2_texto_armado", index=False)
        df_conteo.to_excel(w, sheet_name="3_conteo")
        df_tfidf.to_excel(w, sheet_name="4_tfidf")
        df_idf.to_excel(w, sheet_name="5_explicacion_idf", index=False)
        similitud.to_excel(w, sheet_name="6_parecido_entre_correos")

    print(f"\nExcel generado: {config.RUTA_DEMO_EXCEL}")
    print("Abrelo y compara la hoja 3 (conteo) con la hoja 4 (tfidf).")
    if not es_real:
        print("\nOJO: esto fue con correos de ejemplo. Corre 01_extraer_enviados.py")
        print("para generar la base cruda con tus correos de verdad.")


if __name__ == "__main__":
    main()
