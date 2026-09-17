# -*- coding: utf-8 -*-
"""
Created on Sat Feb 14 10:19:02 2026

@author: USUARIO
"""

import os
from adjuntos.ocr import escanear_y_categorizar_un_archivo
import time
from correo_automation.funciones_bot_continuo import guardar_en_historial
import pandas as pd
from datetime import datetime

# --- FASE 3: DESCARGAR ADJUNTOS Y ANALIZAR ---

# Creamos una carpeta para los PDF si no existe
folder_temp = "adjuntos/adjuntos_temp"

if not os.path.exists(folder_temp):
    os.makedirs(folder_temp)

reporte_final = []

def descargar_leer_y_categorizar_adjuntos(cuenta_objetivo, emails):
    print(f"3. Iniciando análisis de {len(emails)} correos...")

    #cada 5 minutos revise la bandeja de entrada (al revisar se excluyen los que ya se leyeron)
    fecha_y_hora_inicio_revision = datetime.now()
    #Convertimos a lista fija para que el movimiento no rompa el bucle
    lista_emails = list(emails) 
    conteo_movidos = 0
    
    #breakpoint()
    for idx, m in enumerate(lista_emails, 1):
        #breakpoint()
        diferencia  = datetime.now() - fecha_y_hora_inicio_revision
        minutos_en_revision = diferencia.total_seconds() / 60

        if (minutos_en_revision > 10):
            print("Ya pasaron 10 minutos, se revisa de nuevo")
            break
        # Contador de progreso para tu hermano
        if idx % 10 == 0: 
            print(f"📊 Progreso: {idx}/{len(lista_emails)} correos analizados...")

        print(f"  -> Procesando correo asunto: {m.subject}")

        # Guardar info en la lista del reporte
        reporte_final.append({
                        "Email_ID": m.ConversationID,
                        "Archivo": "N/A",
                        "Categoria": "Otros / No leído",
                        "Pagina_Deteccion": 0                     
        })
                    
        #Guardar la info de que se pasó por ese mail
        #La función mantiene solo el más reciente
        guardar_en_historial(pd.DataFrame(reporte_final))
            
        if m.Attachments.Count > 0:
            for adjunto in m.Attachments:
                # Solo nos interesan PDFs
                if adjunto.FileName.lower().endswith(".pdf"):
                    ruta_descarga = os.path.join(os.getcwd(), folder_temp, adjunto.FileName)
                    print(f"   -> Procesando correo Asunto: {m.Subject}")     
                    print(f"   -> Descargando: {adjunto.FileName}")
                    adjunto.SaveAsFile(ruta_descarga)
                    
                    #breakpoint()
                    # Ahora recibimos UN solo objeto (el diccionario)
                    resultado_ocr = escanear_y_categorizar_un_archivo(ruta_descarga)

                    #Eliminar el adjunto descargado después de ocr
                    # 2. ELIMINAR EL ARCHIVO (Clean up)
                    try:
                        if os.path.exists(ruta_descarga):
                            os.remove(ruta_descarga)
                            print(f"🗑️ Archivo adjunto temporal eliminado: {adjunto.FileName}")
                    except Exception as e:
                        print(f"⚠️ No se pudo eliminar el archivo: {e}")
                        
                    # Extraemos los datos del diccionario de forma segura
                    categoria = resultado_ocr.get("categoria", "Error")
                    pagina_hallazgo = resultado_ocr.get("pagina", "N/A")

                    print(f"🎯 Resultado: {categoria} (Encontrado en pág. {pagina_hallazgo})")

                    # Guardar info en la lista del reporte
                    reporte_final.append({
                        "Email_ID": m.ConversationID,
                        "Archivo": adjunto.FileName,
                        "Categoria": categoria,
                        "Pagina_Deteccion": pagina_hallazgo                       
                    })
                    
                    #breakpoint()
                    #Guardar la info de que se leyó ese correo
                    #La función mantiene solo el más reciente
                    guardar_en_historial(pd.DataFrame(reporte_final))

                    #Si ya categorizó el correo pasar al siguiente correo
                    if (categoria != "Otros / No identificado"):
                        
                        #mover el correo
                        try:
                            m.Move(cuenta_objetivo.Folders.Item(categoria))                                                
                            print(f"📦 Movido: {m.Subject} -> {categoria}")
                            # Pausa tras mover (el servidor de correo es lento)
                            time.sleep(1.5)
                            conteo_movidos += 1                        
                        except Exception as e_move:
                                print(f"⚠️ No se pudo mover '{m.Subject}': {e_move}")

                        #pasar al siguiente correo
                        break

    print(f"✅ Correos movidos con éxito, total correos {len(lista_emails)}, cantidad correos leídos {len(reporte_final)} | Movidos: {conteo_movidos}")
    
    return reporte_final