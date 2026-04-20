# Guía de Instalación y Uso del Procesador de Libros

Este script (`separador_mvp.py`) permite procesar un documento de Word (`.docx`) y convertirlo en archivos Markdown para un proyecto de Quarto. Durante el proceso, el script extrae imágenes nativas, captura ecuaciones directamente en la línea de texto y conserva el formato principal.

## Requisitos Previos
1. Tener instalado **Python 3.8 o superior** en tu computadora.
2. Contar con el archivo de Word original en la misma carpeta que el script. Por defecto, el script buscará el archivo `LGRD_CAPITULOS_V16.docx`, pero puedes cambiar este nombre directamente en las últimas líneas del código `separador_mvp.py`.

## Paso 1: Crear el Entorno Virtual (Recomendado)

Para evitar conflictos con otras librerías instaladas en tu computador, se recomienda crear un entorno virtual para este script. Abre una terminal (o consola de comandos) en la carpeta donde tienes los archivos descargados y ejecuta:

**En Windows (Símbolo del sistema o PowerShell):**
```cmd
python -m venv venv
venv\Scripts\activate
```

**En Mac / Linux:**
```bash
python3 -m venv venv
source venv/bin/activate
```
*(Sabrás que funcionó si al inicio de la línea en tu terminal aparece la palabra `(venv)`)*

## Paso 2: Instalar Dependencias

Con el entorno virtual ya activado, instala la única librería externa requerida utilizando el archivo `requirements.txt`:

```bash
pip install -r requirements.txt
```

## Paso 3: Ejecutar el Script

Verifica que el archivo `.docx` se encuentre en la misma carpeta. Para iniciar el procesamiento, ejecuta:

```bash
python separador_mvp.py
```

## ¿Qué sucede después?

Al finalizar, el script creará automáticamente una carpeta llamada `proyecto_libro_quarto`. Dentro de ella encontrarás:
- El archivo `_quarto.yml` configurado con la estructura del libro.
- Un archivo `.qmd` por cada capítulo, lista de revisores, índices, etc.
- Una carpeta `media` que contiene todas las imágenes extraídas del documento original.

Si cuentas con [Quarto](https://quarto.org/docs/get-started/) instalado en tu equipo, puedes ingresar a esa carpeta generada y ejecutar:
```bash
cd proyecto_libro_quarto
quarto preview
```
Para previsualizar cómo se verá tu libro.
