# -*- coding: utf-8 -*-
"""
Created on Wed Feb 11 08:02:11 2026

@author: USUARIO
"""
from __future__ import annotations
from .params import EmailParams
from datetime import datetime, timedelta

def get_last_messages(cuenta_folder, params: EmailParams, carpeta_a_leer = "Bandeja de Entrada"):
    
    inbox = cuenta_folder.Folders(carpeta_a_leer)
    #inbox = cuenta_folder.Folders.Item(5)#Folders("Bandeja de Entrada")

    for i, f in enumerate(cuenta_folder.Folders, 1):
        print(f"Índice {i}: '{f.Name}'")
    print("---------------------------------------------------\n")
    # Filtro: Solo correos de los últimos 2 días para ir rápido
    #hace_dos_dias = (datetime.now() - timedelta(days=2)).strftime("%d/%m/%Y %H:%M %p")
    #filtro_hoy = (datetime.now() - timedelta(days=10)).strftime("%d/%m/%Y %H:%M %p")
    #breakpoint()
    hoy_medianoche = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    #hoy_medianoche = (hoy_medianoche - timedelta(days=1))
    filtro_hoy = hoy_medianoche.strftime("%d/%m/%Y %H:%M %p")

    #Y filtrar solo los no ledídos (los leídos significa ya Migue los gestionó)
    filtro = f"[ReceivedTime] >= '{filtro_hoy}' AND [UnRead] = True"
    
    # Restrict es la clave para que no se quede pegado
    messages = inbox.Items.Restrict(filtro)
    messages.Sort("[ReceivedTime]", True)

    total = messages.Count
    #if (params.iniciar_desde == 1):
    #    total = min(params.iniciar_desde + params.max_emails - 1, messages.Count)
    #else:
    #    total = params.iniciar_desde + params.max_emails - 1


    result = []

    for i in range(params.iniciar_desde, total):
        result.append(messages.Item(i))
    
    print(f"Hay {len(result)} correos de hoy y sin leer en la carpeta '{carpeta_a_leer}'")

    return result

def get_sent_messages(cuenta_folder, params: EmailParams):
    sent_folder = cuenta_folder.Folders("Elementos enviados")
    #sent_folder = cuenta_folder.Folders.Item(6)


    # Filtro: Solo correos de los últimos 2 días para ir rápido
    hace_dos_dias = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y %H:%M %p")
    filtro = f"[SentOn] >= '{hace_dos_dias}'"
    
    # Restrict es la clave para que no se quede pegado
    messages = sent_folder.Items.Restrict(filtro)
    messages.Sort("[SentOn]", True)
    
    # Traemos un número razonable para comparar (ej. el doble de lo que leíste en Inbox)
    total = min(params.max_emails * 2, messages.Count)
    return [messages.Item(i) for i in range(1, total + 1)]