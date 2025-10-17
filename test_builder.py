"""
Script de prueba rápida para DocumentBuilder.
Genera un documento de ejemplo con todas las características básicas.
"""

from word import DocumentBuilder
import pandas as pd


def test_basico():
    """Test básico de todas las funcionalidades."""
    print("=" * 60)
    print("TEST DE DOCUMENTBUILDER")
    print("=" * 60)

    # Crear datos de prueba
    df_ejemplo = pd.DataFrame({
        'Columna A': [1, 2, 3],
        'Columna B': ['X', 'Y', 'Z'],
        'Columna C': [10.5, 20.3, 15.8]
    })

    # Construir documento
    print("\n1. Inicializando builder...")
    builder = DocumentBuilder()

    print("2. Agregando contenido...")
    (builder
        .indice()
        .advertencia_actualizacion()
        .salto_pagina()

        .titulo("Documento de Prueba", 1)
        .parrafo("Este documento fue generado automáticamente por DocumentBuilder para verificar su funcionamiento.")

        .titulo("Sección 1: Texto", 2)
        .parrafo("Este es un párrafo de ejemplo que demuestra el formato de texto.")
        .parrafo("Este es un segundo párrafo para mostrar el espaciado.")

        .titulo("Sección 2: Listas", 2)
        .parrafo("Ejemplo de lista con viñetas:")
        .vinetas([
            "Primera viñeta",
            "Segunda viñeta",
            "Tercera viñeta"
        ])

        .titulo("Sección 3: Tablas", 2)
        .parrafo("Tabla de ejemplo con datos:")
        .tabla(df_ejemplo, titulo="Tabla 1: Datos de Ejemplo")

        .salto_pagina()
        .titulo("Sección 4: Conclusiones", 2)
        .parrafo("El DocumentBuilder funciona correctamente. Todas las operaciones se completaron sin errores.")
        .vinetas([
            "Títulos jerárquicos funcionan",
            "Párrafos con formato correcto",
            "Tablas insertadas correctamente",
            "Encadenamiento de métodos operativo"
        ])
    )

    print("3. Mostrando historial de operaciones...")
    builder.mostrar_historial()

    print("4. Guardando documento...")
    builder.guardar("test_builder_output.docx")

    print("\n" + "=" * 60)
    print("TEST COMPLETADO EXITOSAMENTE")
    print("Archivo generado: test_builder_output.docx")
    print("=" * 60)


if __name__ == "__main__":
    test_basico()
