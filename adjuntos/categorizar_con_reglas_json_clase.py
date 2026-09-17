# -*- coding: utf-8 -*-
"""
Created on Wed Feb 21 08:02:11 2026

@author: Juan Carlos Herrera Burbano
"""

#import json

#class ProcesadorDocumentos:
#    def __init__(self, ruta_json="C:\Users\USUARIO\Juan Carlos\Gestión de Correo - Migue\GIT\configuracion_carpetas.json"):
#        with open(ruta_json, 'r', encoding='utf-8') as f:
#            self.reglas = json.load(f)
import os
import json

class ProcesadorDocumentos:
    def __init__(self, ruta_json=None):
        # Si no se pasa una ruta específica, calculamos la ruta automática
        if ruta_json is None:
            # os.getcwd() toma la carpeta desde donde se ejecuta el comando (la raíz de GIT)
            ruta_json = os.path.join(os.getcwd(), "configuracion_carpetas.json")
            
        print(f"Buscando configuración en: {ruta_json}") # Para validar que apunte bien
        
        with open(ruta_json, 'r', encoding='utf-8') as f:
            self.reglas = json.load(f)

    def categorizar(self, texto_pagina):
        # Aquí SÍ puedes usar self.reglas
        # 2. Iterar sobre cada categoría definida en el JSON
        for categoria in self.reglas:
            cumple_todas = True
            
            # 3. Validar cada regla (regla_1, regla_2, etc.)
            # Todas las reglas deben cumplirse (AND)
            for nombre_regla, palabras_de_las_reglas in categoria["palabras_clave"].items():

                #es mejor usar conjuntos set para no usar in
                #solo se harán 5 iteraciones de palabras_de_las_reglas en lugar de 500                  
                #palabras que pueden haber en  texto_pagina
                if not any(p.lower() in texto_pagina for p in palabras_de_las_reglas):
                    cumple_todas = False
                    break # Si una regla no se cumple, pasamos a la siguiente categoría

            # 4. Si cumplió todas las reglas de esta categoría, devolvemos el resultado
            if cumple_todas:
                return {
                    "categoria": categoria["nombre_carpeta"]
                }
                
        return {"categoria": "Otros / No identificado", "pagina": None}

# Creamos la instancia AQUÍ MISMO una sola vez
motor_categorizacion = ProcesadorDocumentos() 