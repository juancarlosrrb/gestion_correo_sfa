# -*- coding: utf-8 -*-
"""
Created on Wed Feb 21 08:02:11 2026

@author: USUARIO
"""

#Definir la funciÃ³n que escanea los pdfs y los categoriza

import easyocr
from pdf2image import convert_from_path, pdfinfo_from_path
import numpy as np
import os
from adjuntos.categorizar_con_reglas_json_clase import motor_categorizacion#ProcesadorDocumentos   
import fitz

# Inicializamos el lector una sola vez
reader = easyocr.Reader(['es'])

def escanear_y_categorizar_un_archivo(ruta_pdf):
    # 1. Definimos la ruta donde lo instalaste
    #POPPLER_PATH = r'C:\Program Files\Poppler\poppler-24.08.0\Library\bin'
    POPPLER_PATH = r'C:\poppler\Library\bin'

    try:
        # Obtenemos el número total de páginas sin cargar el archivo completo
        info = pdfinfo_from_path(ruta_pdf, poppler_path=POPPLER_PATH)
        total_paginas = info["Pages"]
        
        print(f"Analizando {total_paginas} páginas de {os.path.basename(ruta_pdf)}...")

        # ========================================================
        # 1. INTENTO RÁPIDO: EXTRAER TEXTO NATIVO
        # ========================================================
        doc = fitz.open(ruta_pdf)
        texto_completo = ""
        # Leemos solo las primeras 2 páginas (o las que necesites)
        for pagina in doc.pages(0, 2):
            texto_completo += pagina.get_text()
        doc.close()
        
        # Verificamos si el PDF tiene texto real (y no es solo una imagen escaneada)
        if len(texto_completo.strip()) > 50: # Si hay más de 50 caracteres, asumimos que se leyó bien
            print("✅ Texto nativo detectado, categorizando...")
            resultado_rapido = motor_categorizacion.categorizar(texto_completo.lower())
            
            # Le agregamos la página para mantener la estructura de tu diccionario
            if resultado_rapido["categoria"] != "Otros / No identificado":
                resultado_rapido["pagina"] = 1 # Asignamos 1 por defecto al leer nativo
            else:
                resultado_rapido["pagina"] = None
                
            return resultado_rapido
        
        # ========================================================
        # 2. SI EL TEXTO ES MUY CORTO (O VACÍO), HACEMOS OCR
        # ========================================================
        else:
            print("⚠️ PDF sin texto indexado, iniciando OCR...")
            
            texto_pagina = "" # Inicializamos la variable por si acaso
            
            # Iteramos página por página
            for i in range(1, total_paginas + 1):
                print(f"Escaneando página {i}...")
                
                if (i > 2):
                    print(f"Son muchas páginas, guardando que se leyó y pasando al siguiente adjunto o correo. Página {i}")
                    return {"categoria": "Otros / No identificado", "pagina": None}
                
                # Convertimos SOLO la página actual a imagen
                paginas = convert_from_path(ruta_pdf, first_page=i, last_page=i, dpi=300, poppler_path=POPPLER_PATH)
                if not paginas: continue
                
                imagen_np = np.array(paginas[0])
                
                # EasyOCR lee la página
                resultados_texto = reader.readtext(imagen_np, detail=0, paragraph=True)
                texto_pagina_i = " ".join(resultados_texto).lower()
                            
                #Combinar con texto páginas anteriores
                if i == 1:
                    texto_pagina = texto_pagina_i
                elif (i > 1):
                    texto_pagina = texto_pagina + " " + texto_pagina_i
                
                #Eliminar este texto porque ensucia la lectura
                texto_pagina = texto_pagina.replace("tutelas_fomag@fiduprevisora.com.co", "")

                # Lógica de categorización rápida, anterior:

                #inicio1 borrar después de pruebas
                #if any(palabra in texto_pagina for palabra in ["fomag", "fondo nacional de prestaciones sociales del magisterio", "fnpsm", "f.n.p.s.m", "fondo de prestaciones del magisterio"]):
                    
                #    if any(palabra in texto_pagina for palabra in ["tutela", "acción de tutela"]):
                        
                #        if any(palabra in texto_pagina for palabra in ["desacato"]):
                #            return {
                #                "categoria": "1. TUTELAS FOMAG DESACATO", 
                #                "pagina": i}
                        
                #        else:
                #            return {
                #                "categoria":"2. TUTELAS FOMAG",
                #                "pagina": i
                #                }
                ##fin1 borrar después de pruebas

            #categorizar con reglas json, nuevo:
            resultado_categorizacion = motor_categorizacion.categorizar(texto_pagina)

            # 2. Le agregas la página (esto modifica el diccionario original)
            if resultado_categorizacion["categoria"] != "Otros / No identificado":
                # Como esto corre después del bucle, le asignamos la última página leída 'i'
                resultado_categorizacion["pagina"] = i  
            else:
                resultado_categorizacion["pagina"] = None

            return resultado_categorizacion
        
    except Exception as e:
        print(f"❌ Error al procesar: {str(e)}")
        return {"categoria": "Error OCR/Texto", "pagina": None, "hallazgo": str(e)}