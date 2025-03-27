# ContratoGenerator-

Descripción

Este repositorio contiene el código necesario para automatizar la generación de documentos en formato Word a partir de datos almacenados en un archivo Excel. La automatización permite reemplazar marcadores específicos dentro de una plantilla de Word con la información contenida en el Excel.

Importante: Seguridad de la Información

Por razones de seguridad, el documento Word generado no ha sido subido a este repositorio.
Solo se incluye el código fuente para la automatización. Si deseas generar el informe, debes proporcionar tu propia plantilla en formato .docx, cabe señalar que el documento debe llamarse "Incremento_actividades.docx".

Funcionamiento

El script en Python carga un archivo Excel, extrae los datos y reemplaza las siguientes palabras clave en la plantilla de Word:

NOMBREp → Se reemplaza con el valor de la columna NOMBRE_P

FECHA_ACTUAL → Se reemplaza con la fecha de ejecución

ACTIVIDADp → Se reemplaza con la actividad correspondiente

Estos reemplazos permiten personalizar automáticamente el contenido del documento según los datos proporcionados.

Requisitos

Para ejecutar el código, asegúrate de contar con las siguientes dependencias instaladas:

pip install pandas openpyxl python-docx

Uso

Coloca tu plantilla Word en la carpeta del proyecto.

Asegúrate de tener el archivo Excel con los datos.

Ejecuta el script:

python Incremento_actividades.py

El documento generado estará disponible en la carpeta de salida especificada en el código.

Notas

Asegúrate de que los nombres de las columnas en el archivo Excel coincidan con los utilizados en el código.

Si el script lanza un error de tipo TypeError: replace() argument 2 must be str, not float, verifica que los valores en el Excel no sean nulos o conviértelos a cadena antes del reemplazo.

Revisa la documentación para realizar ajustes en la plantilla de Word según sea necesario.
