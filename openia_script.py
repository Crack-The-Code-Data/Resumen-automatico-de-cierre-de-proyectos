import pandas as pd
import openai
from typing import List, Union, Optional, Dict, Any
import os
from datetime import datetime
import json
import logging
from dotenv import load_dotenv

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Cargar variables de entorno
load_dotenv()

# Configurar la clave API de OpenAI desde variable de entorno
openai.api_key = os.getenv("API_KEY")

if not openai.api_key:
    raise ValueError("API_KEY no encontrada en el archivo .env")

# Lista global para registrar tokens (considera usar un sistema de logging más robusto)
registro_tokens = []

# Diccionario de precios por modelo (USD por 1M tokens)
PRECIOS_MODELOS = {
    'gpt-4.1': {'input': 2.00, 'output': 8.00},
    'gpt-4.1-mini': {'input': 0.40, 'output': 1.60},
    'gpt-4.1-nano': {'input': 0.10, 'output': 0.40},
    'gpt-4.5-preview': {'input': 75.00, 'output': 150.00},
    'gpt-4o': {'input': 2.50, 'output': 10.00},
    'gpt-4o-mini': {'input': 0.15, 'output': 0.60},
    'gpt-4o-mini-realtime-preview': {'input': 0.60, 'output': 2.40},
    'gpt-4o-realtime-preview': {'input': 5.00, 'output': 20.00},
    'gpt-4o-audio-preview': {'input': 2.50, 'output': 10.00},
    'gpt-4o-mini-audio-preview': {'input': 0.15, 'output': 0.60},
    'gpt-4o-search-preview': {'input': 2.50, 'output': 10.00},
    'gpt-4o-mini-search-preview': {'input': 0.15, 'output': 0.60},
    'o1': {'input': 15.00, 'output': 60.00},
    'o1-pro': {'input': 150.00, 'output': 600.00},
    'o3-pro': {'input': 20.00, 'output': 80.00},
    'o3': {'input': 2.00, 'output': 8.00},
    'o3-deep-research': {'input': 10.00, 'output': 40.00},
    'o4-mini': {'input': 1.10, 'output': 4.40},
    'o4-mini-deep-research': {'input': 2.00, 'output': 8.00},
    'o3-mini': {'input': 1.10, 'output': 4.40},
    'o1-mini': {'input': 1.10, 'output': 4.40},
    'codex-mini-latest': {'input': 1.50, 'output': 6.00},
    'computer-use-preview': {'input': 3.00, 'output': 12.00},
    'gpt-image-1': {'input': 5.00, 'output': 1.25},
}

# Prompts del sistema almacenados como constantes
SYSTEM_PROMPT = """
Eres un analista de datos educativos especializado en la redacción de informes técnicos profesionales.

CONTEXTO:
Debes redactar secciones específicas de un informe sobre resultados de proyectos educativos, basándote únicamente en los datos proporcionados.

DIRECTRICES ESTRICTAS:

1. OBJETIVIDAD ABSOLUTA:
   - Describe únicamente lo que muestran los datos, sin interpretaciones causales
   - Evita correlaciones no fundamentadas (ej: "X indica éxito del programa")
   - No atribuyas significado sin evidencia directa
   - Usa lenguaje neutral y descriptivo

2. LENGUAJE PROFESIONAL:
   - Emplea terminología técnica apropiada
   - Redacta en tercera persona
   - Utiliza voz pasiva cuando sea pertinente
   - Mantén un tono formal y académico

3. ESTRUCTURA DE REDACCIÓN:
   - Para introducción: Contextualiza el proyecto y sus objetivos medibles
   - Para resúmenes: Sintetiza hallazgos clave sin valoraciones
   - Para observaciones: Presenta tendencias y patrones identificables
   - Para conclusiones: Resume datos presentados sin extrapolaciones

4. PROHIBICIONES EXPLÍCITAS:
   - NO inferir causalidad sin evidencia
   - NO hacer juicios de valor sobre los datos
   - NO relacionar variables demográficas con éxito/fracaso
   - NO incluir recomendaciones no solicitadas
   - NO usar adjetivos valorativos (exitoso, deficiente, prometedor)

5. FORMATO DE RESPUESTA:
   - Párrafos concisos de 3-5 oraciones
   - Incluye datos específicos cuando sea relevante (porcentajes, cifras)

EJEMPLO DE REDACCIÓN APROPIADA:
Incorrecto: "La alta participación femenina (70%) demuestra el éxito del programa"
Correcto: "La distribución por género muestra una participación del 70% de mujeres y 30% de hombres"

"""

FORMATO_SALIDA_BASICO = "No uses markdown, solo texto plano. No uses titulos, solo párrafos. No uses emojis. No uses saltos de linea. Porcentajes con 1 decimal."

def _detectar_modelo_base(modelo: str) -> str:
    """
    Detecta el modelo base a partir del nombre completo del modelo.

    Args:
        modelo (str): Nombre completo del modelo.

    Returns:
        str: Modelo base encontrado en el diccionario de precios.
    """
    if modelo in PRECIOS_MODELOS:
        return modelo

    # Buscar coincidencia parcial ordenada por longitud (más específico primero)
    modelos_ordenados = sorted(PRECIOS_MODELOS.keys(), key=len, reverse=True)
    for key in modelos_ordenados:
        if modelo.startswith(key):
            return key

    logger.warning(f"Modelo '{modelo}' no encontrado en PRECIOS_MODELOS. Usando precios por defecto.")
    return modelo


def guardar_registro_tokens(archivo: str = "registro_tokens.csv") -> None:
    """
    Guarda el registro de tokens en un archivo CSV.

    Args:
        archivo (str): Nombre del archivo donde guardar el registro.
    """
    try:
        df = pd.DataFrame(registro_tokens)
        # Si el archivo existe, agregamos; si no, creamos uno nuevo
        if os.path.exists(archivo):
            df_existente = pd.read_csv(archivo)
            df = pd.concat([df_existente, df], ignore_index=True)
        df.to_csv(archivo, index=False)
        logger.info(f"Registro de tokens guardado en {archivo}")
    except Exception as e:
        logger.error(f"Error al guardar registro de tokens: {str(e)}")


def call_gpt(
    prompt: str,
    modelo: str = "gpt-4o-mini",
    max_tokens: int = 1500,
    temperature: float = 0.7,
    system_prompt: Optional[str] = None
) -> str:
    """
    Llama a la API de OpenAI. Por defecto usa gpt-4o-mini.

    Args:
        prompt (str): Texto del prompt a enviar.
        modelo (str): Modelo de OpenAI a utilizar.
        max_tokens (int): Máximo de tokens en la respuesta.
        temperature (float): Control de creatividad (0.0-1.0).
        system_prompt (Optional[str]): Prompt del sistema personalizado.

    Returns:
        str: Respuesta del modelo.

    Raises:
        ValueError: Si el prompt está vacío.
    """
    if not prompt or not prompt.strip():
        raise ValueError("El prompt no puede estar vacío")

    if system_prompt is None:
        system_prompt = SYSTEM_PROMPT

    try:
        response = openai.chat.completions.create(
            model=modelo,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            max_tokens=max_tokens,
            temperature=temperature
        )

        result = response.choices[0].message.content.strip()

        # Registrar uso de tokens
        usage = response.usage
        input_tokens = usage.prompt_tokens
        output_tokens = usage.completion_tokens

        modelo_base = _detectar_modelo_base(modelo)
        precios = PRECIOS_MODELOS.get(modelo_base, {'input': 0, 'output': 0})

        cost_usd = (input_tokens * precios['input'] + output_tokens * precios['output']) / 1000000

        registro_tokens.append({
            'fecha_hora': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'modelo': modelo,
            'input_tokens': input_tokens,
            'output_tokens': output_tokens,
            'costo_usd': cost_usd,
        })

        logger.info(f"Tokens usados - Input: {input_tokens}, Output: {output_tokens}, Costo: ${cost_usd:.6f}")

        return result

    except openai.OpenAIError as e:
        logger.error(f"Error en la API de OpenAI: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Error inesperado: {str(e)}")
        raise

def analyze_dataframe(
    df: pd.DataFrame,
    seccion: str = "observacion",
    contexto: str = "",
    tokens: int = 1000,
    modelo: str = "gpt-4o-mini"
) -> str:
    """
    Analiza un DataFrame y genera texto profesional para informes educativos.

    Args:
        df (pd.DataFrame): DataFrame a analizar.
        seccion (str): Tipo de sección del informe ('introduccion', 'resumen', 'observacion', 'conclusion').
        contexto (str): Contexto adicional sobre los datos (nombre del proyecto, período, etc.).
        tokens (int): Máximo de tokens en la respuesta.
        modelo (str): Modelo de OpenAI a utilizar.

    Returns:
        str: Texto generado para el informe.

    Raises:
        ValueError: Si el DataFrame está vacío o la sección no es válida.
    """
    if df.empty:
        raise ValueError("El DataFrame está vacío. No hay datos para analizar.")
    
    secciones_validas = ["introduccion", "resumen", "observacion", "conclusion"]
    if seccion not in secciones_validas:
        raise ValueError(f"Sección debe ser una de: {', '.join(secciones_validas)}")

    # Convertir DataFrame a JSON
    try:
        json_str = df.to_json(orient="records", lines=False, force_ascii=False)
    except Exception as e:
        logger.error(f"Error al convertir DataFrame a JSON: {str(e)}")
        raise

    # Obtener información básica del DataFrame para contexto
    info_basica = f"""
Dimensiones: {df.shape[0]} registros, {df.shape[1]} variables
Columnas: {', '.join(df.columns.tolist())}
"""

    # Instrucciones específicas según la sección
    instrucciones_seccion = {
        "introduccion": "Redacta una introducción contextual basada en los datos disponibles. Menciona el alcance del análisis y las variables principales.",
        "resumen": "Sintetiza los hallazgos principales observables en los datos. Incluye cifras clave y distribuciones relevantes.",
        "observacion": "Describe patrones, tendencias y distribuciones identificables en los datos. Presenta porcentajes y valores cuando sea pertinente.",
        "conclusion": "Resume los datos presentados de manera objetiva, destacando las características principales del conjunto de datos."
    }

    prompt = f"""Eres un analista de datos educativos especializado en la redacción de informes técnicos profesionales.

CONTEXTO DEL PROYECTO:
{contexto if contexto else "Análisis de datos de proyecto educativo"}

INFORMACIÓN DEL DATASET:
{info_basica}

TAREA:
{instrucciones_seccion[seccion]}

DIRECTRICES ESTRICTAS:
1. OBJETIVIDAD: Describe únicamente lo observable en los datos, sin interpretaciones causales
2. NO inferir éxito/fracaso basándote en distribuciones demográficas
3. NO establecer correlaciones no evidenciadas
4. USA lenguaje técnico, formal y neutral
5. INCLUYE valores específicos cuando sea relevante (porcentajes, totales)
6. EVITA adjetivos valorativos (exitoso, deficiente, prometedor)
7. REDACTA en tercera persona

PROHIBICIONES:
- NO uses frases como "esto indica/sugiere/demuestra éxito"
- NO relaciones género, edad u otras variables demográficas con calidad
- NO hagas recomendaciones
- NO uses expresiones especulativas

DATOS A ANALIZAR:
{json_str}

FORMATO DE SALIDA:
- Máximo 3 párrafos
- Cada párrafo de 2-4 oraciones
- Lenguaje profesional y técnico
- Sin bullets ni listas
- Texto corrido y formal

Genera únicamente el texto solicitado para la sección '{seccion}' del informe."""

    return call_gpt(prompt, modelo=modelo, max_tokens=tokens)


def insight_list(
    data_list: List[Union[int, float, str]],
    proyectos: Optional[pd.DataFrame] = None,
    introduccion: str = "",
    tokens: int = 2000,
    modelo: str = "gpt-4o-mini"
) -> Dict[str, Any]:
    """
    Analiza una lista y genera insights estructurados en formato JSON.

    Args:
        data_list (List): Lista de datos a analizar.
        proyectos (Optional[pd.DataFrame]): DataFrame con información de proyectos.
        introduccion (str): Introducción o contexto del análisis.
        tokens (int): Máximo de tokens en la respuesta.
        modelo (str): Modelo de OpenAI a utilizar.

    Returns:
        Dict[str, Any]: Diccionario con los insights estructurados.

    Raises:
        ValueError: Si la lista está vacía o si el JSON retornado es inválido.
    """
    if not data_list:
        raise ValueError("La lista de datos está vacía. No hay datos para analizar.")

    # Convertir lista a string
    list_str = ", ".join(str(item) for item in data_list)

    json_str = ""
    if proyectos is not None and not proyectos.empty:
        try:
            json_str = proyectos.to_json(orient="records", lines=False, force_ascii=False)
        except Exception as e:
            logger.warning(f"Error al convertir proyectos a JSON: {str(e)}")
            json_str = ""

    # Construir prompt base
    base_prompt = f"""Basándote en la siguiente introducción, información de proyectos y conclusiones parciales, genera un resumen estructurado en formato JSON que destaque los principales hallazgos e insights por dimensión o categoría.

Introducción:
{introduccion}

Proyectos:
{json_str}

Conclusiones parciales:
{list_str}

Formato de salida (devuelve solo un JSON válido):
{{
  "Contexto General del Diagnóstico": [
    "Insight 1",
    "Insight 2",
    "Insight 3"
  ],
  "Hallazgos Clave y Correlaciones Relevantes": {{
    "<Nombre de la categoría>": [
      "Insight 1",
      "Insight 2",
      "Insight 3",
      "Implicación: ..."
    ]
  }},
  "Retos Priorizados Identificados": [
    {{
      "Eje": "Nombre del eje",
      "Reto": "Descripción del reto",
      "Relevancia": "Razón por la cual es importante"
    }}
  ],
  "Otras Secciones Relevantes": {{
    "Título de la sección": [
      "Insight 1",
      "Insight 2",
      "Insight 3"
    ]
  }},
  "Relevancia del Programa": [
    "Punto 1 sobre impacto del programa",
    "Punto 2",
    "Punto 3"
  ]
}}

Instrucciones:
- No incluyas ningún texto fuera del JSON, asegúrate de que el json sea válido.
- Si alguna sección no aplica, omítela (no dejes campos vacíos).
- Usa nombres de categoría o sección que surjan naturalmente del análisis.
- Redacta en estilo claro y sintético.
- Las implicaciones deben reflejar posibles líneas de acción o interpretaciones del dato."""

    response = call_gpt(base_prompt, modelo=modelo, max_tokens=tokens)

    # Validar que el JSON retornado sea válido
    try:
        # Intentar extraer JSON del texto (por si incluye markdown)
        json_text = response.strip()
        if json_text.startswith("```json"):
            json_text = json_text.split("```json")[1].split("```")[0].strip()
        elif json_text.startswith("```"):
            json_text = json_text.split("```")[1].split("```")[0].strip()

        resultado = json.loads(json_text)
        logger.info("JSON de insights validado correctamente")
        return resultado
    except json.JSONDecodeError as e:
        logger.error(f"Error al parsear JSON de insights: {str(e)}")
        logger.error(f"Respuesta recibida: {response}")
        # Retornar el texto como fallback
        return {"error": "JSON inválido", "respuesta_original": response}