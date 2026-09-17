# -*- coding: utf-8 -*-
"""
Created on Wed Feb 11 07:56:15 2026

@author: USUARIO
"""

from datetime import datetime, timedelta
import time
from correo_automation.connection import connect_outlook
from correo_automation.reader import get_last_messages
from reporte.generar_reporte import reporte_adjuntos, categorizar_emails
from reenviar_correos.reenviar_correos_de_carpeta import reeenviar_correos
from correo_automation.funciones_bot_continuo import (
    guardar_en_historial, es_horario_laboral, filtrar_correos_nuevos
)
import time

def ejecutar_ciclo_bot(cuenta_objetivo, params):
    """
    Orquestador principal: Prioriza FOMAG y luego Bandeja de Entrada.
    """
    if not es_horario_laboral():
        return False 

    print(f"\n🔔 [{datetime.now().strftime('%H:%M:%S')}] Iniciando revisión...")

    # --- FASE 1: REVISAR FOMAG PRIMERO ---
    params.iniciar_desde = 1
    carpeta_actual = "FOMAG"
    
    try:
        emails_inbox = get_last_messages(cuenta_objetivo, params, carpeta_a_leer=carpeta_actual)
    except:
        # Conectar de nuevo si falla
        session, cuenta_objetivo = connect_outlook(params.email_objetivo)
        emails_inbox = get_last_messages(cuenta_objetivo, params, carpeta_a_leer=carpeta_actual)

    emails_para_procesar = filtrar_correos_nuevos(emails_inbox)

    # --- FASE 2: SI FOMAG ESTÁ AL DÍA, PASAR A BANDEJA DE ENTRADA ---
    if len(emails_para_procesar) == 0:
        print("✅ FOMAG al día. Pasando a revisar Bandeja de Entrada...")
        
        carpeta_actual = "Bandeja de Entrada"
        params.iniciar_desde = 1 # Reiniciar el contador para la nueva carpeta
        
        # Leer la primera página de la Bandeja de Entrada
        emails_inbox = get_last_messages(cuenta_objetivo, params, carpeta_a_leer=carpeta_actual)
        emails_para_procesar = filtrar_correos_nuevos(emails_inbox)

        # --- FASE 3: BÚSQUEDA PROFUNDA (Solo aplica para Bandeja de Entrada) ---
        n_revisiones = 0
        hora_inicio = datetime.now()

        while len(emails_para_procesar) == 0:

            #Esperar 4 minutos
            time.sleep(60*4)

            hora_actual = datetime.now()
            minutos_transcurridos = (hora_actual - hora_inicio).total_seconds() / 60
            
            if minutos_transcurridos >= 5:
                print(f"⚠️ ¡Tiempo límite! ({minutos_transcurridos:.1f} min). Aplicando break.")
                break

            n_revisiones += 1
            # Sumar lo que ya se leyó
            params.iniciar_desde = 1 + (n_revisiones * params.max_emails)
            
            # Volver a leer más atrás
            emails_inbox = get_last_messages(cuenta_objetivo, params, carpeta_a_leer=carpeta_actual)
            emails_para_procesar = filtrar_correos_nuevos(emails_inbox)

            if len(emails_para_procesar) > 0:
                break

    # --- FASE 4: PROCESAR LOS CORREOS ENCONTRADOS ---
    # Llegado a este punto, emails_para_procesar tiene correos de FOMAG *o* de Bandeja de Entrada
    if len(emails_para_procesar) > 0:
        print(f"🚀 Procesando {len(emails_para_procesar)} correos nuevos de la carpeta: {carpeta_actual}...")
        df_categorizados = categorizar_emails(cuenta_objetivo, emails_para_procesar, params)

        if df_categorizados is not None and not df_categorizados.empty:
            # Persistencia y Salida
            reporte_adjuntos(emails_inbox=emails_para_procesar, df_emails_categorizados=df_categorizados)
            #reeenviar_correos(cuenta_objetivo)
            
    else:
        print("🙏 Todo al día en ambas carpetas. Esperando próximo ciclo...")
        
    return True