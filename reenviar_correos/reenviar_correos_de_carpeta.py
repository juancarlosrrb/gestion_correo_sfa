# -*- coding: utf-8 -*-
"""
Created on Wed Feb 25 07:43:12 2026

@author: USUARIO
"""

from .config import cargar_configuracion
import time

# El "Cerebro" que despacha cada tarea

def reeenviar_correos(cuenta_objetivo):
    config_carpetas = cargar_configuracion()
    #breakpoint()
    for config in config_carpetas:
        reenviar_contenido_carpeta(cuenta_objetivo, config)


def reenviar_contenido_carpeta(cuenta_objetivo, config):
    """
    Recibe el objeto de la cuenta y UN diccionario de configuración.
    """
    nombre_origen = config['nombre_carpeta']
    destinatario = config['correo_reenviar']
    nombre_destino = config['carpeta_procesados']

    try:
        # 1. Acceder a las carpetas (Origen y Destino)
        # IMPORTANTE: Outlook necesita el OBJETO de la carpeta para el .Move()
        carpeta_origen = cuenta_objetivo.Folders.Item(nombre_origen)
        carpeta_destino = cuenta_objetivo.Folders.Item(nombre_destino)
        
        items = carpeta_origen.Items
        total = items.Count

        if total == 0:
            print(f"📭 La carpeta '{nombre_origen}' está vacía.")
            return

        print(f"🚀 Iniciando ráfaga: {total} correos de '{nombre_origen}' -> {destinatario}")

        # Iteramos de atrás hacia adelante (indispensable al mover items)
        for i in range(total, 0, -1):
            m = items.Item(i)
            
            try:
                # 2. Reenvío
                reenvio = m.Forward()
                reenvio.To = destinatario
                reenvio.Send()
                # Pausa tras enviar (el servidor de correo es lento)
                time.sleep(0.3)

                # 4. Guardar como leído
                #if m.UnRead:
                #    m.UnRead = False
                #    #Escribe estos cambios:
                #    m.Save()
                #    time.sleep(0.1)

                # 3. Mover (Pasamos el objeto carpeta_destino, no el string)
                try:
                    m.Move(carpeta_destino)
                except:
                    copia_en_destino = m.Copy().Move(carpeta_destino)
    
                    # 2. Verificamos que la copia exista realmente
                    if copia_en_destino:
                        # 3. Solo si la copia tuvo éxito, marcamos como No Leída en destino
                        #copia_en_destino.UnRead = True
                        copia_en_destino.Save()
                        
                        # 4. AHORA SÍ, borramos el original del origen
                        m.Delete() 
                        print(f"✅ Proceso seguro: {copia_en_destino.Subject}...")
                    else:
                        print("⚠️ No se pudo confirmar la creación de la copia.")

                # Pausa tras enviar (el servidor de correo es lento)
                time.sleep(0.5)

                
                print(f"   ✅ [{total-i+1}/{total}] Reenviado y Movido: {m.Subject[:30]}...")
                

            except Exception as e_mail:
                print(f"   ❌ Error con correo {i}: {e_mail}")

        print(f"✨ ¡Carpeta {nombre_origen} completada!")

    except Exception as e:
        print(f"❌ Error crítico accediendo a carpetas: {e} config es: {config}")