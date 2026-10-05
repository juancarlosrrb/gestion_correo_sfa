# -*- coding: utf-8 -*-
"""
Created on Fri Feb 13 05:57:09 2026

@author: USUARIO
"""
import os
import sys
import pandas as pd
from datetime import datetime

# Forzar a Python a reconocer la carpeta raíz 'GIT' como base de búsqueda
ruta_raiz = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ruta_raiz not in sys.path:
    sys.path.append(ruta_raiz)
    
from adjuntos.descargar_adjuntos import descargar_leer_y_categorizar_adjuntos
import time
from correo_automation.connection import connect_outlook
from correo_automation.reader import get_last_messages

def df_procesar_inbox(emails_entrada):
    datos_reporte = []

    for email in emails_entrada:
        try:
            # Extraemos los datos básicos de cada objeto de Outlook
            info = {
                "ID_Conversacion": email.ConversationID,
                "Asunto": email.Subject,
                #"Remitente": email.SenderName,
                "Correo Remitente": email.SenderEmailAddress,
                "Fecha Recibido": str(email.ReceivedTime), # Convertimos a string para evitar líos de zona horaria
                #"Cuerpo (Resumen)": email.Body[:100] + "..." if email.Body else ""
            }
            datos_reporte.append(info)
        except Exception as e:
            print(f"Error al leer un correo: {e}")

    # Creamos un DataFrame (una tabla de datos)
    df = pd.DataFrame(datos_reporte)

    return df


def reporte_bandeja_entrada(emails, carpeta_salida = "reportes_generados"):
    """
    Toma una lista de objetos de correo de Outlook y los exporta a un Excel.
    """

    df = df_procesar_inbox(emails)
    
    # 1. Crear la carpeta si no existe para evitar errores
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)
        

    # Generamos el nombre del archivo con la fecha actual
    nombre_archivo = f"reporte1_correos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    ruta_final = os.path.join(carpeta_salida, nombre_archivo)
    
    # Guardamos a Excel
    df.to_excel(ruta_final, index=False)
    
    print(f"✅ Reporte generado con éxito: {nombre_archivo}")
    

def reporte_comparativo_control(emails_inbox, emails_sent, carpeta_salida = "reportes_generados"):
    # --- 1. Procesar Bandeja de Entrada ---
    df_inbox = df_procesar_inbox(emails_inbox)

    # --- 2. Procesar Enviados ---
    # Limpiamos el asunto para que coincida (quitamos el "RV: " o "FW: ")
    data_sent = []
    for m in emails_sent:
        asunto_limpio = m.Subject.replace("RV: ", "").replace("FW: ", "").strip()
        data_sent.append({
            "ID_Conversacion": m.ConversationID,
            "Asunto_Enviado": m.Subject,
            "Asunto_Limpio_Enviado": asunto_limpio,
            "Fecha_Envío": str(m.SentOn),
            "Enviado": "SÍ"
        })
    df_sent = pd.DataFrame(data_sent).drop_duplicates(subset=['ID_Conversacion'])

    print(df_sent)
    print(df_inbox)
    if (df_inbox.shape[0] == 0 or len(data_sent) == 0):
        print("⚠️ AVISO: No se encontraron correos enviados. No se generará el reporte comparativo.")
        return None # Devolvemos un mensaje para que el main sepa qué pasó
    
    # --- 3. El LEFT JOIN ---
    # Unimos por ID_Conversacion que es lo más preciso en Outlook
    #breakpoint()
    reporte_final = pd.merge(df_inbox, df_sent, on="ID_Conversacion", how="left")
    # Rellenamos los que no tienen coincidencia con "NO"
    reporte_final["Enviado"] = reporte_final["Enviado"].fillna("NO")

    # --- 4. Guardar ---
    #reporte_final.to_excel("reporte2_gestion_legal.xlsx", index=False)
    
    # 1. Crear la carpeta si no existe para evitar errores
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)
        
    # Generamos el nombre del archivo con la fecha actual
    nombre_archivo = f"reporte2_correos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    ruta_final = os.path.join(carpeta_salida, nombre_archivo)
    
    # Guardamos a Excel
    reporte_final.to_excel(ruta_final, index=False)
    
    print(f"✅ Reporte generado con éxito: {nombre_archivo}")
    print("Reporte generado. Revisa la columna 'Enviado' para ver los pendientes.")
    
  
def categorizar_emails(cuenta_objetivo, emails_inbox, params):
    
    #Hacer ocr:
    reporte_final = descargar_leer_y_categorizar_adjuntos(cuenta_objetivo, emails_inbox)

    # --- FASE DE REPORTE (Usando Pandas) ---
    df = pd.DataFrame(reporte_final)
    
    # Generamos el nombre del archivo con la fecha actual
    nombre_archivo = f"Reporte2_categorizacion_OCR_correos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    ruta_final = os.path.join("reportes_generados", nombre_archivo)

    df.to_excel(ruta_final, index=False)

    print("📊 Reporte OCR Excel generado con éxito.")
            
    return df

def reporte_adjuntos(emails_inbox, df_emails_categorizados):
    
    df_emails_inbox = df_procesar_inbox(emails_inbox)
    
    # 2. Definimos las columnas que ESPERAMOS del categorizado
    # Ajusta estos nombres según lo que devuelva tu proceso de OCR
    columnas_esperadas = ['Email_ID', 'categoria', 'pagina'] 

    # 3. Blindaje: Si viene None o vacío, creamos un cascarón con las columnas
    if df_emails_categorizados is None or df_emails_categorizados.empty:
        print("⚠️ Advertencia: No hay correos categorizados. Creando reporte vacío.")
        df_emails_categorizados = pd.DataFrame(columns=columnas_esperadas)

    #breakpoint()
    df_resultado = pd.merge(
        df_emails_inbox, 
        df_emails_categorizados, 
        left_on= 'ID_Conversacion',
        right_on='Email_ID', 
        how='left')
    
    # Generamos el nombre del archivo con la fecha actual
    nombre_archivo = f"Reporte2_categorizacion_OCR_correos_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    ruta_final = os.path.join("reportes_generados", nombre_archivo)

    df_resultado.to_excel(ruta_final, index=False)

    print("📊 Reporte OCR Excel generado con éxito.")
            
    return df_resultado