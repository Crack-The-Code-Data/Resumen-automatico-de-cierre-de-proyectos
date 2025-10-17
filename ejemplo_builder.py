"""
Ejemplos de uso de DocumentBuilder para generar documentos Word.

Este archivo muestra diferentes formas de usar la clase DocumentBuilder
para crear documentos profesionales de manera fluida y eficiente.
"""

from word import DocumentBuilder
from docx.shared import Inches, RGBColor
import pandas as pd


def ejemplo_basico():
    """Ejemplo básico: Documento simple con encadenamiento."""
    print("\n=== EJEMPLO 1: Documento Básico ===")

    builder = DocumentBuilder()
    builder.titulo("Mi Primer Reporte", 1) \
           .parrafo("Este es un documento creado con DocumentBuilder.") \
           .titulo("Introducción", 2) \
           .parrafo("El patrón Builder facilita la creación de documentos complejos.") \
           .vinetas([
               "Encadenamiento de métodos",
               "Configuración centralizada",
               "Historial de operaciones"
           ]) \
           .guardar("ejemplo_basico.docx")


def ejemplo_con_indice():
    """Ejemplo con índice automático y múltiples secciones."""
    print("\n=== EJEMPLO 2: Documento con Índice ===")

    builder = DocumentBuilder()
    builder.indice() \
           .advertencia_actualizacion() \
           .salto_pagina() \
           .titulo("Capítulo 1: Introducción", 1) \
           .parrafo("Primera sección del documento.") \
           .titulo("Sección 1.1", 2) \
           .parrafo("Subsección con más detalles.") \
           .titulo("Sección 1.2", 2) \
           .parrafo("Otra subsección importante.") \
           .salto_pagina() \
           .titulo("Capítulo 2: Desarrollo", 1) \
           .parrafo("Contenido del segundo capítulo.") \
           .titulo("Conclusiones", 1) \
           .parrafo("Resumen final del documento.") \
           .guardar("ejemplo_indice.docx")


def ejemplo_con_tabla():
    """Ejemplo con tablas desde DataFrames de pandas."""
    print("\n=== EJEMPLO 3: Documento con Tablas ===")

    # Crear datos de ejemplo
    df_ventas = pd.DataFrame({
        'Producto': ['Laptop', 'Mouse', 'Teclado', 'Monitor'],
        'Cantidad': [45, 120, 89, 34],
        'Precio': [850, 25, 45, 320],
        'Total': [38250, 3000, 4005, 10880]
    })

    df_empleados = pd.DataFrame({
        'Departamento': ['IT', 'IT', 'Ventas', 'Ventas', 'HR'],
        'Nombre': ['Juan', 'María', 'Carlos', 'Ana', 'Luis'],
        'Salario': [5000, 5500, 4000, 4200, 3800]
    })

    builder = DocumentBuilder()
    builder.titulo("Reporte de Ventas y Personal", 1) \
           .titulo("Ventas del Trimestre", 2) \
           .parrafo("A continuación se muestra el detalle de ventas por producto.") \
           .tabla(df_ventas, titulo="Tabla de Ventas") \
           .salto_pagina() \
           .titulo("Personal por Departamento", 2) \
           .parrafo("Lista de empleados organizados por departamento.") \
           .tabla(df_empleados, titulo="Empleados", con_merge=True, group_cols=['Departamento']) \
           .guardar("ejemplo_tablas.docx")


def ejemplo_configuracion_personalizada():
    """Ejemplo con configuración personalizada de estilos."""
    print("\n=== EJEMPLO 4: Configuración Personalizada ===")

    # Configuración custom
    config_custom = {
        'fuente_titulo': 'Arial',
        'fuente_texto': 'Calibri',
        'color_titulo': RGBColor(0x00, 0x5A, 0x9C),  # Azul corporativo
        'margin_top': Inches(0.5),
        'margin_bottom': Inches(0.5),
    }

    builder = DocumentBuilder(config=config_custom)
    builder.titulo("Reporte Corporativo", 1) \
           .parrafo("Este documento usa estilos personalizados.") \
           .titulo("Objetivos", 2) \
           .vinetas([
               "Usar fuentes corporativas",
               "Aplicar colores de marca",
               "Mantener consistencia visual"
           ]) \
           .guardar("ejemplo_custom.docx")


def ejemplo_completo():
    """Ejemplo completo con todas las características."""
    print("\n=== EJEMPLO 5: Documento Completo ===")

    # Datos de ejemplo
    df_metricas = pd.DataFrame({
        'Métrica': ['Usuarios Activos', 'Ingresos', 'Satisfacción', 'Retención'],
        'Valor Actual': [15420, '$125,000', '4.5/5', '87%'],
        'Objetivo': [20000, '$150,000', '4.7/5', '90%'],
        'Estado': ['En progreso', 'En progreso', 'Alcanzado', 'En progreso']
    })

    builder = DocumentBuilder()

    # Construir documento completo
    (builder
        .indice()
        .advertencia_actualizacion()
        .salto_pagina()

        .titulo("Reporte Ejecutivo Q1 2025", 1)
        .parrafo("Este documento presenta un resumen completo de los resultados del primer trimestre.")

        .titulo("1. Resumen Ejecutivo", 2)
        .parrafo("Durante el primer trimestre se lograron avances significativos en los siguientes áreas:")
        .vinetas([
            "Incremento del 23% en usuarios activos",
            "Mejora en indicadores de satisfacción",
            "Lanzamiento de nuevas funcionalidades"
        ])

        .titulo("2. Métricas Principales", 2)
        .parrafo("Las siguientes métricas muestran el desempeño general del periodo:")
        .tabla(df_metricas, titulo="Tabla 1: Indicadores Clave de Desempeño")

        .titulo("3. Análisis Detallado", 2)
        .titulo("3.1 Crecimiento de Usuarios", 3)
        .parrafo("El crecimiento de usuarios se debió principalmente a campañas de marketing digital.")

        .titulo("3.2 Ingresos", 3)
        .parrafo("Los ingresos mostraron una tendencia positiva durante todo el trimestre.")

        .salto_pagina()

        .titulo("4. Conclusiones y Recomendaciones", 2)
        .parrafo("Basado en los resultados del trimestre, se recomienda:")
        .vinetas([
            "Continuar con las estrategias actuales de adquisición",
            "Invertir en mejoras de producto",
            "Expandir el equipo de soporte al cliente"
        ], nivel=1)

        .titulo("5. Próximos Pasos", 2)
        .vinetas([
            "Definir objetivos para Q2",
            "Asignar presupuesto adicional",
            "Programar revisión mensual de métricas"
        ], nivel=1)

        .mostrar_historial()
        .guardar("ejemplo_completo.docx")
    )


def ejemplo_historial():
    """Ejemplo mostrando el uso del historial."""
    print("\n=== EJEMPLO 6: Uso del Historial ===")

    builder = DocumentBuilder()
    builder.titulo("Documento de Prueba", 1) \
           .parrafo("Primer párrafo.") \
           .parrafo("Segundo párrafo.") \
           .vinetas(["Item 1", "Item 2"])

    # Obtener y mostrar historial
    historial = builder.obtener_historial()
    print("\nOperaciones realizadas:")
    for i, op in enumerate(historial, 1):
        print(f"  {i}. {op}")

    builder.guardar("ejemplo_historial.docx")


def ejemplo_manejo_errores():
    """Ejemplo mostrando manejo de errores con DataFrames vacíos."""
    print("\n=== EJEMPLO 7: Manejo de Errores ===")

    df_vacio = pd.DataFrame()

    builder = DocumentBuilder()
    builder.titulo("Reporte con Tabla Vacía", 1) \
           .parrafo("Intentando insertar una tabla vacía...") \
           .tabla(df_vacio, titulo="Esta tabla está vacía") \
           .parrafo("El builder maneja automáticamente el error.") \
           .guardar("ejemplo_errores.docx")


if __name__ == "__main__":
    print("=" * 60)
    print("EJEMPLOS DE USO DE DOCUMENTBUILDER")
    print("=" * 60)

    # Ejecutar todos los ejemplos
    ejemplo_basico()
    ejemplo_con_indice()
    ejemplo_con_tabla()
    ejemplo_configuracion_personalizada()
    ejemplo_completo()
    ejemplo_historial()
    ejemplo_manejo_errores()

    print("\n" + "=" * 60)
    print("✓ Todos los ejemplos se ejecutaron correctamente")
    print("✓ Revisa los archivos .docx generados en el directorio actual")
    print("=" * 60)
