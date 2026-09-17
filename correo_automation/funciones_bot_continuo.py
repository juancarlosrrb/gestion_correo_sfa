# -*- coding: utf-8 -*-
"""
Created on Wed Feb 11 07:56:15 2026

@author: USUARIO
"""
import os
import pandas as pd
from datetime import datetime

PATH_HISTORIAL = "historial_procesados.parquet"

def cargar_historial():
    """Carga los IDs ya procesados desde el archivo Parquet."""
    if os.path.exists(PATH_HISTORIAL):
        try:
            return pd.read_parquet(PATH_HISTORIAL)
        except Exception as e:
            print(f"⚠️ Error leyendo historial, creando uno nuevo: {e}")
    
    return pd.DataFrame(columns=['Email_ID', 'Categoria', 'fecha_proceso'])

def guardar_en_historial(df_nuevo):
    """Agrega los nuevos correos procesados al archivo de memoria."""
    if df_nuevo is None or df_nuevo.empty:
        return
    
    historial = cargar_historial()
    df_nuevo['fecha_proceso'] = datetime.now()
    
    # Concatenamos y eliminamos duplicados por MessageID
    actualizado = pd.concat([historial, df_nuevo], ignore_index=True)
    actualizado['Email_ID'] = actualizado['Email_ID'].astype(str)
    actualizado = actualizado.sort_values('fecha_proceso')
    actualizado = actualizado.drop_duplicates(subset=['Email_ID'], keep='last')
    
    actualizado.to_parquet(PATH_HISTORIAL, index=False)
    print(f"💾 Historial actualizado: {len(actualizado)} registros totales.")

def es_horario_laboral():
    """Verifica si el bot debe seguir trabajando (Corte 5:00 PM)."""
    ahora = datetime.now()
    print("Es horario laboral true")
    #if ahora.hour >= 17:
    #    print(f"🌆 [{ahora.strftime('%H:%M')}] Fin de jornada. ¡Hasta mañana!")
    #    return False
    return True

def filtrar_correos_nuevos(emails_todos):
    """Compara la bandeja de entrada con el historial y devuelve solo lo no leído."""
    #breakpoint()
    historial = cargar_historial()
    # Convertimos a set para que la búsqueda sea ultra rápida (O(1))
    ids_procesados = set(historial['Email_ID'].astype(str).tolist())
    
    nuevos = [e for e in emails_todos if str(e.ConversationID) not in ids_procesados]
    return nuevos