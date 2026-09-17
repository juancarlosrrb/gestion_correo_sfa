# -*- coding: utf-8 -*-
"""
Created on Wed Feb 11 08:01:17 2026

@author: USUARIO
"""

import win32com.client
from datetime import datetime, timedelta
import os
import time

def connect_outlook(email_objetivo):

    ##cuando cambien correo:
    # 1. Forzamos el cierre de cualquier sesión fantasma
    #os.system("taskkill /f /im outlook.exe 2>nul")
    
    # Damos 1 segundo para que Windows termine de cerrar el proceso
    #time.sleep(1)
    #outlook = win32com.client.Dispatch("Outlook.Application")
    #namespace = outlook.GetNamespace("MAPI")
    #namespace.Logon("", "", True, True)

    try:
        # Intenta conectar con la instancia abierta
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except:
        # Si no hay instancia, la crea de cero
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            print("❌ No se pudo iniciar Outlook. Intenta abrir Outlook manualmente primero.")
            raise e
        
    namespace = outlook.GetNamespace("MAPI")
    #namespace.Logon("", "", True, True)
    for i, folder in enumerate(namespace.Folders):
        print(f"Revisando Cuenta {i}: -->{folder.Name}<--")
    
        ##Vamos por cada cuenta y cada carpeta para ver si hay info
        #for carpeta in folder.Folders():
        #    nombre_carpeta = carpeta.Name.lower()
            
        #Ir por cada cuenta y ver cantidad de enviados últimos dos días (para pruebas)
        try:
            sent_folder = folder.Folders("Elementos enviados")
        
            #Filtro: Solo correos de los últimos 2 días para ir rápido
            hace_dos_dias = (datetime.now() - timedelta(days=2)).strftime("%d/%m/%Y %H:%M %p")
            filtro = f"[SentOn] >= '{hace_dos_dias}'"
        
            # Restrict es la clave para que no se quede pegado
            messages = sent_folder.Items.Restrict(filtro)
            cantidad_enviados = messages.Count
            print(f"Cuenta {i}, nombre cuenta {folder.Name}. Cantidad de enviados dos días antes: {cantidad_enviados}")
        except:
            print("No encontró Elementos enviados")
            
        if folder.Name == email_objetivo:
            print(f"Cuenta {i} escogida: {folder.Name}")
            return namespace, folder # Retornamos la cuenta específica

    print(f"❌ Error: No se encontró la cuenta objetivo '{email_objetivo}', ver main.py definición params")   
    return namespace, None
        
        
        