# Pipeline de modelo para clasificar correos

Cinco scripts independientes. Cada uno deja un archivo en `datos_modelo/`
que el siguiente lee. Si uno falla, los anteriores no se pierden.

## Dónde poner esto

Copia la carpeta `modelo_correos/`, los 6 scripts y `config_areas.json`
dentro de tu proyecto `GIT/`, al lado de `main.py`:

```
GIT/
├── main.py                        (lo que ya tenías)
├── correo_automation/             (lo que ya tenías — se reutiliza)
├── adjuntos/                      (lo que ya tenías)
├── reenviar_correos/              (lo que ya tenías)│
├── modelo_correos/                ← NUEVO (librería compartida)
├── 01_extraer_enviados.py         ← NUEVO
├── 02_vectorizar.py
├── 02b_demo_texto_a_numeros.py
├── 03_entrenar_modelo.py
├── 04_predecir_bandeja.py
├── 05_mover_correos.py
├── config_areas.json
└── datos_modelo/                  (se crea solo)
```

Instalación: `py -m pip install -r requisitos_modelo.txt` (o línea por línea).

## Orden de uso

```bash
# 1. Construir la base histórica (la primera vez, varias corridas)
py 01_extraer_enviados.py --cantidad 100      # empieza chiquito para ver que sirve
py 01_extraer_enviados.py --cantidad 2000     # cuando confíes, dale duro

# 2. Entender cómo el texto se vuelve número (opcional pero recomendado)
py 02b_demo_texto_a_numeros.py

# 3. Vectorizar y entrenar
py 02_vectorizar.py
py 03_entrenar_modelo.py

# 4. Predecir sobre la bandeja de entrada
py 04_predecir_bandeja.py --cantidad 100

# 5. Revisar el Excel a mano y después mover
py 05_mover_correos.py                        # simula, no mueve nada
py 05_mover_correos.py --ejecutar             # mueve de verdad
```

## Antes de correr el 04

Edita `config_areas.json`. La llave es el correo al que tu hermano reenvía;
el valor dice a qué carpeta de Outlook se mueve el original. Las carpetas
deben existir con el mismo nombre exacto en Outlook (el script 05 no crea
carpetas a propósito: no quiero que un typo llene el buzón de carpetas).

## Archivos que se generan

| Archivo | Lo produce | Para qué |
|---|---|---|
| `base_de_datos_0historica_cruda_modelo.parquet` | 01 | 10 correos completos, sin recortar, para aprender |
| `base_de_datos_1historica_modelo.parquet` | 01 | La base histórica que crece cada día |
| `demo_texto_a_numeros.xlsx` | 02b | Texto → conteo → TF-IDF, hoja por hoja |
| `vectorizador_tfidf.joblib` | 02 | El diccionario palabra→columna. **Obligatorio al predecir** |
| `matriz_X.npz` / `etiquetas_y.parquet` | 02 | Los números y las respuestas |
| `modelo_clasificador.joblib` | 03 | El modelo entrenado |
| `resultados_modelo.xlsx` | 03 | Qué tan bien predice, y en qué se equivocó |
| `predicciones_bandeja.xlsx` | 04 | Lo que propone mover (lo revisas tú) |
| `bitacora_movidos.xlsx` | 05 | Qué se movió, cuándo y si falló |

## Tres cosas que hay que saber

**1. El vectorizador y el modelo van en pareja.**
Si vuelves a correr el 02, tienes que volver a correr el 03. Si no, el
modelo viejo recibe columnas que no reconoce y predice basura con toda
la confianza del mundo.

**2. El ID que uso no es `ConversationID`.**
Tu bot actual usa `ConversationID`, que identifica el hilo, no el correo.
Si a tu hermano le responden tres veces el mismo asunto, los tres correos
comparten ese ID y el historial creería que ya los procesó. Uso el
Message-ID de internet, que es único por correo.

**3. El 05 no mueve nada sin `--ejecutar`.**
Es a propósito. Un error de clasificación con 2.000 correos al día es un
desastre que se arregla a mano, uno por uno.

## Qué hacer con los resultados del paso 03

Mira la tabla de umbrales que imprime. Te dice algo como: "con umbral 0.90
mueve el 70% de los correos y de esos acierta el 98%". Ese es el número que
decide la operación: ese 70% se mueve solo, el 30% restante lo sigue viendo
tu hermano. No hace falta que el modelo sea perfecto para que le quite la
mayor parte del trabajo.

La hoja `3_matriz_confusion` te dice qué áreas se confunden entre sí. Si
área1 y área2 se confunden mucho, casi siempre es porque la regla real de
tu hermano usa algo que no está en el texto (quién manda el correo, la
fecha, el número de radicado). Eso se arregla agregando ese dato como
columna, no cambiando de algoritmo.

## Cuando la base crezca

Con 2.000 correos/día el parquet llega a ~1 GB en unos 3 meses. Cuando
eso moleste:

- Parte la base por mes: `base_de_datos_1historica_modelo_2026-09.parquet`.
- O deja de guardar `texto_adjuntos` una vez confirmes que el asunto solo
  ya da buen acierto (pruébalo: baja `MAX_CHARS_ADJUNTOS` a 0 en
  `config.py`, corre 02 y 03, y compara el acierto de test).

Casi siempre el asunto y los nombres de los adjuntos cargan el 90% de la
señal. Vale la pena medirlo antes de asumir que necesitas leerlo todo.
