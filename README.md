# Generador de Documentos

Una aplicación de escritorio escrita en Python (usando `tkinter`, `docxtpl` y `docx2pdf`) que permite automatizar la creación de oficios y documentos a partir de plantillas de Word (`.docx`).

## Características
- **Interfaz gráfica simple**: Selecciona tu plantilla y llena los campos requeridos.
- **Detección automática de etiquetas**: El programa lee tu `.docx` y crea dinámicamente un formulario basado en las etiquetas `{{ etiqueta }}`.
- **Variables automáticas y autoincremento**: Soporte para un archivo `.json` del mismo nombre que la plantilla. Permite definir variables ocultas y números de oficio que se incrementan automáticamente en cada uso (con soporte para relleno de ceros, ej. `001`).
- **Exportación a PDF**: Opción de guardar el resultado final como Word (`.docx`) o convertirlo automáticamente a PDF.

## Requisitos
- Python 3.x
- Microsoft Word (requerido únicamente si deseas exportar a PDF)

## Instalación
Instala las dependencias necesarias ejecutando:
```bash
pip install -r requirements.txt
```

## Uso
Ejecuta la aplicación:
```bash
python app.py
```
