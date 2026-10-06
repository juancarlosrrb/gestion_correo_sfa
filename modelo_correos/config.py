# -*- coding: utf-8 -*-
"""
Configuracion central del pipeline de modelo.
Todo lo que se cambia seguido vive aca, no regado en los scripts.
"""
import os

# ----------------------------------------------------------------------
# CUENTA Y CARPETAS DE OUTLOOK
# ----------------------------------------------------------------------
#EMAIL_OBJETIVO = "notjudicial5@fiduprevisora.com.co"
EMAIL_OBJETIVO = "juan_cherrerab@soy.sena.edu.co"


CARPETA_ENVIADOS = "Elementos enviados"
CARPETA_ENTRADA = "Bandeja de Entrada"

# ----------------------------------------------------------------------
# RUTAS DE ARCHIVOS (todo se guarda en la carpeta datos_modelo/)
# ----------------------------------------------------------------------
CARPETA_DATOS = "datos_modelo"

BASE_CRUDA = os.path.join(CARPETA_DATOS, "base_de_datos_0historica_cruda_modelo.parquet")
BASE_MODELO = os.path.join(CARPETA_DATOS, "base_de_datos_1historica_modelo.parquet")

# Salidas del paso 02 (vectorizacion)
RUTA_VECTORIZADOR = os.path.join(CARPETA_DATOS, "vectorizador_tfidf.joblib")
RUTA_MATRIZ_X = os.path.join(CARPETA_DATOS, "matriz_X.npz")
RUTA_ETIQUETAS_Y = os.path.join(CARPETA_DATOS, "etiquetas_y.parquet")
RUTA_DEMO_EXCEL = os.path.join(CARPETA_DATOS, "demo_texto_a_numeros.xlsx")

# Salidas del paso 03 (entrenamiento)
RUTA_MODELO = os.path.join(CARPETA_DATOS, "modelo_clasificador.joblib")
RUTA_RESULTADOS_EXCEL = os.path.join(CARPETA_DATOS, "resultados_modelo.xlsx")
RUTA_RESULTADOS_JSON = os.path.join(CARPETA_DATOS, "resultados_modelo.json")

# Salidas del paso 04 (prediccion sobre bandeja de entrada)
RUTA_PREDICCIONES_EXCEL = os.path.join(CARPETA_DATOS, "predicciones_bandeja.xlsx")
RUTA_PREDICCIONES_PARQUET = os.path.join(CARPETA_DATOS, "predicciones_bandeja.parquet")

# Salida del paso 05 (bitacora de correos movidos)
RUTA_BITACORA_MOVIDOS = os.path.join(CARPETA_DATOS, "bitacora_movidos.xlsx")

# Mapeo destinatario -> carpeta destino
RUTA_CONFIG_AREAS = "config_areas.json"

# Carpeta temporal para bajar adjuntos antes de leerlos
CARPETA_TEMP_ADJUNTOS = os.path.join(CARPETA_DATOS, "adjuntos_temp")

# ----------------------------------------------------------------------
# LIMITES DE TAMANO
# Con 2.000 correos/dia la base crece rapido. Estos topes son lo que
# mantiene el parquet manejable sin perder senal para el modelo.
# ----------------------------------------------------------------------
MAX_CHARS_CUERPO = 5000          # el asunto y el arranque del cuerpo cargan casi toda la senal
MAX_CHARS_ADJUNTOS = 16000        # texto extraido de TODOS los adjuntos de un correo, sumado
MAX_PAGINAS_PDF = 3              # solo las primeras paginas de cada PDF
MAX_MB_ADJUNTO = 15              # adjuntos mas pesados que esto se saltan (solo se guarda el nombre)

# Cada cuantos correos se graba el parquet (proteccion contra caidas)
CHECKPOINT_CADA = 25

# Cuantos correos completos y sin recortar se guardan en la base cruda (la de aprender)
N_CORREOS_BASE_CRUDA = 10

# ----------------------------------------------------------------------
# PARAMETROS DEL MODELO
# ----------------------------------------------------------------------
PROPORCION_TEST = 0.2            # 20% de los correos se reservan para evaluar
SEMILLA = 42                     # para que los resultados se puedan repetir igual
MIN_CORREOS_POR_CLASE = 5        # destinatarios con menos correos que esto se descartan del entrenamiento

# El asunto pesa mas que el cuerpo: se repite N veces al armar el texto
PESO_ASUNTO = 3

# Palabras vacias del espanol (las que no aportan nada para clasificar)
STOP_WORDS_ES = [
    "a", "al", "algo", "algunas", "algunos", "ante", "antes", "como", "con", "contra",
    "cual", "cuando", "de", "del", "desde", "donde", "durante", "e", "el", "ella",
    "ellas", "ellos", "en", "entre", "era", "es", "esa", "ese", "eso", "esta",
    "estas", "este", "esto", "estos", "ha", "han", "hasta", "hay", "la", "las",
    "le", "les", "lo", "los", "mas", "me", "mi", "mucho", "muy", "no", "nos",
    "o", "otra", "otro", "para", "pero", "poco", "por", "porque", "que", "quien",
    "se", "sea", "ser", "si", "sin", "sobre", "solo", "son", "su", "sus", "también",
    "tambien", "te", "tiene", "tienen", "todo", "todos", "tu", "un", "una", "uno",
    "unos", "y", "ya",
]


def asegurar_carpetas():
    """Crea las carpetas de trabajo si todavia no existen."""
    os.makedirs(CARPETA_DATOS, exist_ok=True)
    os.makedirs(CARPETA_TEMP_ADJUNTOS, exist_ok=True)
