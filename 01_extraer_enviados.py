# -*- coding: utf-8 -*-
"""
PASO 01 - Construir la base historica desde "Elementos enviados".

Que hace:
  - Se conecta a Outlook y abre la carpeta de enviados.
  - Recorre los correos del mas nuevo al mas viejo.
  - Se SALTA los que ya estan en la base (por email_id), asi que manana
    puedes correrlo otra vez y solo lee lo nuevo.
  - De cada correo saca: a quien se lo envio (TARGET), asunto, cuerpo,
    nombres de adjuntos y el texto de adentro de los adjuntos.
  - Graba en base_de_datos_1historica_modelo.parquet cada 25 correos,
    para que si se cae Outlook no se pierda el trabajo.
  - La primera vez tambien crea base_de_datos_0historica_cruda_modelo.parquet
    con 10 correos SIN recortar, para que veas como se ve el dato completo.

Uso:
    py 01_extraer_enviados.py
    py 01_extraer_enviados.py --cantidad 500
    py 01_extraer_enviados.py --cantidad 2000 --ocr
"""
import argparse
import sys
import time

sys.path.insert(0, ".")

from correo_automation.connection import connect_outlook

from modelo_correos.almacenamiento import (
    cargar_base,
    guardar_base,
    ids_ya_leidos,
)

from modelo_correos import config
from modelo_correos.progreso import barra
from modelo_correos.utils_outlook import extraer_correo

def leer_argumentos():
    p = argparse.ArgumentParser(description="Extrae correos enviados a la base historica.")
    p.add_argument("--cantidad", type=int, default=2000,
                   help="Cuantos correos NUEVOS leer en esta corrida (por defecto 100).")
    p.add_argument("--carpeta", type=str, default=config.CARPETA_ENVIADOS,
                   help='Carpeta a leer (por defecto "Elementos enviados").')
    p.add_argument("--cuenta", type=str, default=config.EMAIL_OBJETIVO,
                   help="Cuenta de Outlook a usar.")
    p.add_argument("--ocr", action="store_true",
                   help="Hacer OCR a los PDF escaneados. LENTO: usalo solo en lotes chicos.")
    p.add_argument("--sin-adjuntos", action="store_true",
                   help="No abrir adjuntos, solo guardar los nombres. Mucho mas rapido.")
    p.add_argument("--max-revisar", type=int, default=0,
                   help="Tope de correos a revisar buscando nuevos. 0 = automatico (cantidad x 20).")
    return p.parse_args()


def main():
    args = leer_argumentos()
    config.asegurar_carpetas()

    print("=" * 70)
    print("PASO 01 - EXTRACCION DE CORREOS ENVIADOS")
    print("=" * 70)

    # ---------------------------------------------------------------
    # 1. Cargar lo que ya se leyo antes
    # ---------------------------------------------------------------
    df_base = cargar_base(config.BASE_MODELO)
    ids_vistos = ids_ya_leidos(df_base)
    print(f"Correos ya registrados en la base: {len(ids_vistos)}")

    # ---------------------------------------------------------------
    # 2. Conectar a Outlook
    # ---------------------------------------------------------------
    print(f"\nConectando a Outlook (cuenta: {args.cuenta})...")
    session, cuenta = connect_outlook(args.cuenta)
    if cuenta is None:
        print("ERROR: no se encontro la cuenta. Revisa el nombre exacto en Outlook.")
        return

    try:
        carpeta = cuenta.Folders(args.carpeta)
    except Exception as e:
        print(f"ERROR: no se pudo abrir la carpeta '{args.carpeta}': {e}")
        print("Carpetas disponibles en esta cuenta:")
        for i, f in enumerate(cuenta.Folders, 1):
            print(f"   {i}. {f.Name}")
        return

    items = carpeta.Items
    # Del mas reciente al mas viejo: asi lo nuevo entra primero
    items.Sort("[SentOn]", True)
    total_en_carpeta = items.Count
    print(f"Carpeta '{args.carpeta}': {total_en_carpeta} correos en total.")

    max_revisar = args.max_revisar if args.max_revisar > 0 else args.cantidad * 20
    max_revisar = min(max_revisar, total_en_carpeta)

    # ---------------------------------------------------------------
    # 3. Recorrer buscando correos NUEVOS
    # ---------------------------------------------------------------
    print(f"\nObjetivo: {args.cantidad} correos nuevos "
          f"(revisando hasta {max_revisar} para encontrarlos).")
    if args.ocr:
        print("OCR ACTIVADO: esto va lento, paciencia.")
    print()

    filas_nuevas = []
    revisados = 0
    saltados = 0
    errores = 0
    inicio = time.time()

    pbar = barra(args.cantidad, "Leyendo correos")

    try:
        for indice in range(1, max_revisar + 1):
            if len(filas_nuevas) >= args.cantidad:
                break

            revisados += 1

            try:
                m = items.Item(indice)
            except Exception:
                errores += 1
                continue

            # Solo correos (una cita o un aviso de entrega no sirven)
            try:
                if str(m.Class) != "43":
                    continue
            except Exception:
                pass

            try:
                fila = extraer_correo(
                    m,
                    usar_ocr=args.ocr,
                    recortar=False,
                    leer_adjuntos=not args.sin_adjuntos,
                )
            except Exception as e:
                errores += 1
                print(f"\n   ! Error leyendo el correo #{indice}: {e}")
                continue

            if not fila["email_id"]:
                errores += 1
                continue

            # Ya lo teniamos -> no cuenta, seguimos buscando
            if fila["email_id"] in ids_vistos:
                saltados += 1
                continue

            if not fila["destinatario"]:
                # Sin destinatario no hay target, no sirve para entrenar
                saltados += 1
                continue

            ids_vistos.add(fila["email_id"])
            filas_nuevas.append(fila)

            pbar.update(1)
            try:
                pbar.set_postfix_str(fila["destinatario"][:32])
            except Exception:
                pass

            # Checkpoint: grabar de a poquitos
            if len(filas_nuevas) % config.CHECKPOINT_CADA == 0:
                df_base = guardar_base(df_base, filas_nuevas, config.BASE_MODELO)
                filas_nuevas = []
    finally:
        pbar.close()

    # ---------------------------------------------------------------
    # 4. Grabar lo que falte
    # ---------------------------------------------------------------
    if filas_nuevas:
        df_base = guardar_base(df_base, filas_nuevas, config.BASE_MODELO)

    duracion = time.time() - inicio

    print("\n" + "=" * 70)
    print("RESUMEN")
    print("=" * 70)
    print(f"Correos revisados en Outlook : {revisados}")
    print(f"Saltados (ya leidos o sin To): {saltados}")
    print(f"Errores                      : {errores}")
    print(f"Tiempo                       : {duracion/60:.1f} min")
    print(f"Base historica               : {config.BASE_MODELO}")
    print(f"Total acumulado en la base    : {len(df_base)} correos")

    if len(df_base) > 0:
        print("\nCorreos por destinatario (top 15):")
        conteo = df_base["destinatario"].value_counts().head(15)
        for correo, n in conteo.items():
            print(f"   {n:6d}  {correo}")

    # ---------------------------------------------------------------
    # 5. Base cruda de 10 correos (la de aprender)
    # ---------------------------------------------------------------
    crear_base_cruda(cuenta, args)


def crear_base_cruda(cuenta, args):
    """
    Guarda 10 correos completos, SIN recortar nada.
    Sirve para que compares a ojo el dato crudo contra lo que el
    modelo termina viendo (numeros).
    """
    import os

    if os.path.exists(config.BASE_CRUDA):
        print(f"\nLa base cruda ya existe ({config.BASE_CRUDA}), no se toca.")
        return

    print(f"\nCreando base cruda de {config.N_CORREOS_BASE_CRUDA} correos (completos, sin recortar)...")

    try:
        carpeta = cuenta.Folders(args.carpeta)
        items = carpeta.Items
        items.Sort("[SentOn]", True)
    except Exception as e:
        print(f"! No se pudo crear la base cruda: {e}")
        return

    filas = []
    indice = 1
    with barra(config.N_CORREOS_BASE_CRUDA, "Base cruda") as pbar:
        while len(filas) < config.N_CORREOS_BASE_CRUDA and indice <= items.Count:
            try:
                m = items.Item(indice)
                fila = extraer_correo(m, usar_ocr=args.ocr, recortar=False, leer_adjuntos=True)
                if fila["email_id"] and fila["destinatario"]:
                    filas.append(fila)
                    pbar.update(1)
            except Exception:
                pass
            indice += 1

    if filas:
        guardar_base(None, filas, config.BASE_CRUDA)
        print(f"Base cruda lista: {config.BASE_CRUDA} ({len(filas)} correos)")
        print("Miralo con:  py 02b_demo_texto_a_numeros.py")


if __name__ == "__main__":
    main()
