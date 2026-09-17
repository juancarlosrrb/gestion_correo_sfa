# -*- coding: utf-8 -*-
"""
Created on Thu Feb 26 07:24:11 2026

@author: USUARIO
"""

import json

def cargar_configuracion():
    try:
        with open('configuracion_carpetas.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"❌ Error cargando config.json: {e}")
        return []

# Cargar al inicio
#CONFIG_CARPETAS = cargar_configuracion()