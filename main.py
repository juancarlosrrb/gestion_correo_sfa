# -*- coding: utf-8 -*-
"""
Created on Wed Feb 11 07:56:15 2026

@author: USUARIO
"""
import time
from correo_automation.orquestador import ejecutar_ciclo_bot
from correo_automation.params import EmailParams
from correo_automation.connection import connect_outlook

# 2. Conexión
params = EmailParams(email_objetivo="notjudicial5@fiduprevisora.com.co", 
                     #email_objetivo="juan_cherrerab@soy.sena.edu.co",
                     max_emails=20,
                     iniciar_desde =  1)

session, cuenta_objetivo = connect_outlook(params.email_objetivo)

if __name__ == "__main__":
    print("🤖 Bot de Gestión iniciado. Modo: Orquestación Modular.")
    
    while True:
        try:
            # Llamamos al orquestador que está en la carpeta correo_automation
            if not ejecutar_ciclo_bot(cuenta_objetivo, params):
                break 
        except Exception as e:
            print(f"❌ Error crítico en el ciclo: {e}")
            break
        
        # Pausa de 1 minuto
        #time.sleep(60*1) 

    print("👋 Bot apagado exitosamente.")