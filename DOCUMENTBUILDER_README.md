# DocumentBuilder - Guía de Uso

## Índice
- [¿Qué es DocumentBuilder?](#qué-es-documentbuilder)
- [Instalación y Requisitos](#instalación-y-requisitos)
- [Uso Básico](#uso-básico)
- [Métodos Disponibles](#métodos-disponibles)
- [Ejemplos Avanzados](#ejemplos-avanzados)
- [Configuración Personalizada](#configuración-personalizada)
- [Comparación con Funciones Tradicionales](#comparación-con-funciones-tradicionales)

---

## ¿Qué es DocumentBuilder?

`DocumentBuilder` es una clase que implementa el **patrón Builder** para facilitar la creación de documentos Word (.docx) de manera fluida y expresiva.

### Ventajas principales:

✅ **Encadenamiento de métodos** - Escribe código más limpio y legible
✅ **Configuración centralizada** - Cambia estilos en un solo lugar
✅ **Historial de operaciones** - Debugging y auditoría simplificados
✅ **Manejo de errores integrado** - Captura automática de excepciones
✅ **Reutilización** - Crea plantillas de configuración personalizadas

---

## Instalación y Requisitos

### Dependencias

```bash
pip install python-docx pandas
```

### Importación

```python
from word import DocumentBuilder
from docx.shared import Inches, RGBColor
```

---

## Uso Básico

### Ejemplo Simple

```python
from word import DocumentBuilder

# Crear builder
builder = DocumentBuilder()

# Construir documento (encadenamiento)
builder.titulo("Mi Reporte", 1) \
       .parrafo("Este es el contenido principal.") \
       .vinetas(["Punto 1", "Punto 2", "Punto 3"]) \
       .guardar("mi_reporte.docx")
```

### Ejemplo sin Encadenamiento

```python
builder = DocumentBuilder()
builder.titulo("Mi Reporte", 1)
builder.parrafo("Primer párrafo.")
builder.parrafo("Segundo párrafo.")
builder.guardar("reporte.docx")
```

---

## Métodos Disponibles

### 1. `titulo(texto, nivel=1)`
Agrega un título con formato jerárquico.

```python
builder.titulo("Título Principal", 1)      # H1 - Centrado, azul marino
builder.titulo("Subtítulo", 2)              # H2 - Izquierda, gris oscuro
builder.titulo("Sección Menor", 3)          # H3 - Itálica, gris
```

**Parámetros:**
- `texto` (str): Texto del título
- `nivel` (int): Nivel jerárquico (1, 2 o 3)

**Retorna:** `self` (para encadenamiento)

---

### 2. `parrafo(texto)`
Agrega un párrafo justificado.

```python
builder.parrafo("Este es un párrafo de texto normal.")
```

**Parámetros:**
- `texto` (str): Contenido del párrafo

**Retorna:** `self`

---

### 3. `tabla(df, titulo=None, con_merge=False, group_cols=None)`
Inserta una tabla desde un DataFrame de pandas.

```python
import pandas as pd

df = pd.DataFrame({
    'Producto': ['Laptop', 'Mouse', 'Teclado'],
    'Precio': [850, 25, 45],
    'Stock': [10, 150, 75]
})

# Tabla simple
builder.tabla(df, titulo="Inventario")

# Tabla con merge de celdas
df_empleados = pd.DataFrame({
    'Depto': ['IT', 'IT', 'Ventas', 'Ventas'],
    'Nombre': ['Juan', 'Ana', 'Carlos', 'María']
})
builder.tabla(df_empleados, titulo="Personal",
              con_merge=True, group_cols=['Depto'])
```

**Parámetros:**
- `df` (DataFrame): Datos de la tabla
- `titulo` (str, opcional): Título sobre la tabla
- `con_merge` (bool): Si usar merge de celdas
- `group_cols` (list): Columnas para agrupar (con merge)

**Retorna:** `self`

---

### 4. `figura(figura, titulo=None, pie=None)`
Inserta una figura de matplotlib.

```python
import matplotlib.pyplot as plt

fig, ax = plt.subplots()
ax.plot([1, 2, 3], [1, 4, 9])
ax.set_title('Gráfico de ejemplo')

builder.figura(fig,
               titulo="Figura 1: Crecimiento",
               pie="Fuente: Datos internos 2025")
```

**Parámetros:**
- `figura`: Objeto Figure de matplotlib
- `titulo` (str, opcional): Título sobre la figura
- `pie` (str, opcional): Pie de figura

**Retorna:** `self`

---

### 5. `vinetas(items, nivel=1, espacio_antes=Pt(4), espacio_despues=Pt(4))`
Agrega lista con viñetas (guiones).

```python
builder.vinetas([
    "Primera viñeta",
    "Segunda viñeta",
    "Tercera viñeta"
])

# Con indentación
builder.vinetas(["Subitem 1", "Subitem 2"], nivel=2)
```

**Parámetros:**
- `items` (list): Lista de strings
- `nivel` (int): Nivel de indentación
- `espacio_antes` (Pt): Espaciado superior
- `espacio_despues` (Pt): Espaciado inferior

**Retorna:** `self`

---

### 6. `salto_pagina()`
Inserta un salto de página.

```python
builder.titulo("Capítulo 1", 1) \
       .parrafo("Contenido...") \
       .salto_pagina() \
       .titulo("Capítulo 2", 1)
```

**Retorna:** `self`

---

### 7. `indice(titulo="Índice")`
Inserta tabla de contenidos automática (TOC).

```python
builder.indice()  # Título por defecto: "Índice"
builder.indice(titulo="Tabla de Contenidos")
```

⚠️ **Nota:** El índice debe actualizarse manualmente en Word:
1. Abrir documento
2. Clic derecho sobre el índice
3. Seleccionar "Actualizar campos"

**Retorna:** `self`

---

### 8. `advertencia_actualizacion()`
Agrega mensaje recordatorio sobre actualización de campos.

```python
builder.indice() \
       .advertencia_actualizacion()  # Mensaje en rojo
```

**Retorna:** `self`

---

### 9. `numerar_titulos()`
Numera automáticamente todos los títulos (1., 1.1, 1.1.1).

```python
builder.titulo("Introducción", 1) \
       .titulo("Contexto", 2) \
       .titulo("Objetivos", 2) \
       .titulo("Desarrollo", 1) \
       .numerar_titulos()  # Aplica numeración
```

**Resultado:**
```
1. INTRODUCCIÓN
  1.1 Contexto
  1.2 Objetivos
2. DESARROLLO
```

**Retorna:** `self`

---

### 10. `guardar(ruta, verbose=True)`
Guarda el documento.

```python
builder.guardar("reporte.docx")
builder.guardar("reporte.docx", verbose=False)  # Sin mensajes
```

**Parámetros:**
- `ruta` (str): Ruta del archivo .docx
- `verbose` (bool): Si mostrar información del guardado

**Retorna:** `None`

---

### 11. `obtener_historial()`
Retorna lista con todas las operaciones realizadas.

```python
historial = builder.obtener_historial()
for operacion in historial:
    print(operacion)
```

**Retorna:** `List[str]`

---

### 12. `mostrar_historial()`
Imprime el historial en consola con formato.

```python
builder.titulo("Ejemplo", 1) \
       .parrafo("Texto...") \
       .mostrar_historial()
```

**Salida:**
```
=== Historial de Operaciones ===
 1. Título nivel 1: Ejemplo
 2. Párrafo: Texto...

Total: 2 operaciones
```

**Retorna:** `self`

---

### 13. `obtener_documento()`
Retorna el objeto `Document` interno para manipulación avanzada.

```python
doc = builder.obtener_documento()
# Ahora puedes usar funciones de python-docx directamente
```

⚠️ **Advertencia:** Modificar directamente puede afectar el historial.

**Retorna:** `Document`

---

## Ejemplos Avanzados

### Documento con Índice Completo

```python
builder = DocumentBuilder()

(builder
    .indice()
    .advertencia_actualizacion()
    .salto_pagina()

    .titulo("Capítulo 1: Introducción", 1)
    .parrafo("Contenido de introducción...")
    .titulo("Antecedentes", 2)
    .parrafo("Contexto histórico...")

    .salto_pagina()
    .titulo("Capítulo 2: Metodología", 1)
    .parrafo("Descripción de métodos...")

    .salto_pagina()
    .titulo("Capítulo 3: Resultados", 1)
    .tabla(df_resultados, "Tabla 1: Datos Principales")

    .guardar("informe_completo.docx")
)
```

### Reporte con Múltiples Tablas

```python
import pandas as pd

df_q1 = pd.DataFrame({'Mes': ['Ene', 'Feb', 'Mar'], 'Ventas': [100, 150, 200]})
df_q2 = pd.DataFrame({'Mes': ['Abr', 'May', 'Jun'], 'Ventas': [180, 220, 250]})

builder = DocumentBuilder()
builder.titulo("Reporte Anual", 1) \
       .titulo("Q1 - Primer Trimestre", 2) \
       .tabla(df_q1, "Ventas Q1") \
       .titulo("Q2 - Segundo Trimestre", 2) \
       .tabla(df_q2, "Ventas Q2") \
       .guardar("reporte_ventas.docx")
```

### Documento con Figuras

```python
import matplotlib.pyplot as plt

# Crear gráficos
fig1, ax1 = plt.subplots()
ax1.bar(['A', 'B', 'C'], [10, 20, 15])

fig2, ax2 = plt.subplots()
ax2.plot([1, 2, 3], [2, 4, 3])

# Insertar en documento
builder = DocumentBuilder()
builder.titulo("Análisis Visual", 1) \
       .figura(fig1, titulo="Figura 1", pie="Comparación de categorías") \
       .figura(fig2, titulo="Figura 2", pie="Tendencia temporal") \
       .guardar("reporte_visual.docx")
```

---

## Configuración Personalizada

### Crear Configuración Custom

```python
from docx.shared import Inches, RGBColor, Pt

config_corporativa = {
    # Márgenes estrechos
    'margin_top': Inches(0.5),
    'margin_bottom': Inches(0.5),
    'margin_left': Inches(0.75),
    'margin_right': Inches(0.75),

    # Fuentes
    'fuente_titulo': 'Arial',
    'fuente_texto': 'Calibri',

    # Colores corporativos
    'color_titulo': RGBColor(0x00, 0x5A, 0x9C),      # Azul corporativo
    'color_subtitulo': RGBColor(0x44, 0x44, 0x44),   # Gris oscuro

    # Tamaños de fuente
    'size_titulo_1': Pt(16),
    'size_texto': Pt(10),

    # Tablas
    'ancho_tabla_default': 6.5,
}

# Usar configuración
builder = DocumentBuilder(config=config_corporativa)
builder.titulo("Reporte Corporativo", 1) \
       .parrafo("Este documento usa estilos corporativos.") \
       .guardar("reporte_corp.docx")
```

### Configuraciones Predefinidas

```python
# Configuración minimalista
config_minimal = {
    'fuente_titulo': 'Helvetica',
    'fuente_texto': 'Helvetica',
    'color_titulo': RGBColor(0x00, 0x00, 0x00),  # Negro
    'size_titulo_1': Pt(12),
    'size_texto': Pt(9),
}

# Configuración académica
config_academica = {
    'fuente_titulo': 'Times New Roman',
    'fuente_texto': 'Times New Roman',
    'size_titulo_1': Pt(14),
    'size_texto': Pt(12),
    'margin_top': Inches(1),
    'margin_bottom': Inches(1),
}

# Usar según necesidad
builder_minimal = DocumentBuilder(config=config_minimal)
builder_academico = DocumentBuilder(config=config_academica)
```

---

## Comparación con Funciones Tradicionales

### ❌ Método Tradicional (Funciones)

```python
from word import crear_documento_a4, agregar_titulo, agregar_parrafo, \
                 insertar_tabla, agregar_viñetas, guardar_documento

doc = crear_documento_a4()
agregar_titulo(doc, "Reporte", 1)
agregar_parrafo(doc, "Introducción...")
agregar_titulo(doc, "Sección 1", 2)
agregar_parrafo(doc, "Contenido...")
agregar_viñetas(doc, ["Item 1", "Item 2"])
insertar_tabla(doc, df, "Tabla 1")
guardar_documento(doc, "reporte.docx")
```

**Problemas:**
- Repetir `doc` en cada llamada
- Difícil rastrear qué se hizo
- No hay configuración centralizada
- Código verboso

---

### ✅ Método Moderno (DocumentBuilder)

```python
from word import DocumentBuilder

builder = DocumentBuilder()
builder.titulo("Reporte", 1) \
       .parrafo("Introducción...") \
       .titulo("Sección 1", 2) \
       .parrafo("Contenido...") \
       .vinetas(["Item 1", "Item 2"]) \
       .tabla(df, "Tabla 1") \
       .guardar("reporte.docx")
```

**Ventajas:**
- ✅ Código más limpio y legible
- ✅ Encadenamiento fluido
- ✅ Historial automático
- ✅ Configuración reutilizable
- ✅ Manejo de errores integrado

---

## Mejores Prácticas

### 1. Usar Encadenamiento para Bloques Relacionados

```python
# ✅ Bueno
builder.titulo("Sección", 2) \
       .parrafo("Texto relacionado") \
       .vinetas(["A", "B", "C"])

# ❌ Menos legible
builder.titulo("Sección", 2)
builder.parrafo("Texto relacionado")
builder.vinetas(["A", "B", "C"])
```

### 2. Usar Variables para Datos Complejos

```python
# ✅ Bueno
items_conclusion = [
    "Resultado 1 alcanzado",
    "Resultado 2 pendiente",
    "Resultado 3 superado"
]
builder.vinetas(items_conclusion)

# ❌ Difícil de mantener
builder.vinetas(["Resultado 1 alcanzado", "Resultado 2 pendiente", ...])
```

### 3. Activar Historial en Desarrollo

```python
builder = DocumentBuilder()
# ... operaciones ...
builder.mostrar_historial()  # Ver qué se hizo
builder.guardar("test.docx")
```

### 4. Reutilizar Configuraciones

```python
# Archivo: configs.py
CONFIG_REPORTE_MENSUAL = { ... }
CONFIG_REPORTE_ANUAL = { ... }

# Usar en tu código
from configs import CONFIG_REPORTE_MENSUAL
builder = DocumentBuilder(config=CONFIG_REPORTE_MENSUAL)
```

---

## Ejecución de Ejemplos

### Desde línea de comandos

```bash
# Usar funciones tradicionales
python word.py mi_documento.docx

# Usar DocumentBuilder
python word.py mi_documento.docx --builder

# Ejecutar todos los ejemplos
python ejemplo_builder.py
```

---

## Troubleshooting

### El índice no se actualiza
**Solución:** Abre el documento en Word, clic derecho en el índice → "Actualizar campos"

### Las tablas se ven desalineadas
**Solución:** Ajusta `ancho_tabla_default` en la configuración:
```python
config = {'ancho_tabla_default': 6.5}
builder = DocumentBuilder(config=config)
```

### Error con DataFrames vacíos
**Solución:** El builder maneja esto automáticamente, pero puedes validar:
```python
if not df.empty:
    builder.tabla(df)
```

### Fuentes no se aplican correctamente
**Solución:** Verifica que la fuente esté instalada en el sistema:
```python
config = {'fuente_titulo': 'Arial'}  # Usar fuentes estándar
```

---

## Recursos Adicionales

- **Documentación python-docx:** https://python-docx.readthedocs.io/
- **Ejemplos completos:** Ver `ejemplo_builder.py`
- **Código fuente:** Ver `word.py` (líneas 438-746)

---

## Licencia y Contribuciones

Este código es parte del proyecto "Resumen automático de cierre de proyectos".
Para contribuir o reportar issues, contacta al equipo de desarrollo.
