# Toolkit Workflow con G.Sheets

## Introducción

Este repositorio contiene un script de Google Apps Script que automatiza el envío de correos electrónicos a partir de datos ingresados en un formulario de Google Sheets. Además, guarda la información en una base de datos interna y limpia el formulario para nuevos ingresos. Este flujo de trabajo es útil para gestionar la recopilación y el procesamiento de datos de manera eficiente, permitiendo a los usuarios enviar información y recibir confirmaciones sin intervención manual.

## Funcionalidades del Código

### Función Principal: `enviar_form()`

1. **Obtención de Datos de Configuración**: Extrae información relevante de la hoja de configuración, como el asunto, destinatarios del correo, cuerpo del mensaje y la configuración del identificador.
2. **Validación de Correos Electrónicos**: Verifica que las direcciones de correo proporcionadas sean válidas.
3. **Recopilación de Datos del Formulario**: Obtiene los datos ingresados en el formulario.
4. **Verificación de Formulario Vacío**: Si el formulario está vacío, muestra una advertencia y detiene el proceso.
5. **Actualización del Identificador**: Genera un nuevo identificador y lo actualiza en la hoja de configuración.
6. **Creación del Cuerpo del Correo**: Construye el cuerpo del correo electrónico en formato HTML, incluyendo una tabla con los datos del formulario.
7. **Preparación de Archivo Adjunto**: Adjunta un archivo al correo si se proporciona una URL válida de Google Drive.
8. **Envío del Correo Electrónico**: Envía el correo con los datos recopilados y el archivo adjunto, si está disponible.
9. **Almacenamiento de Datos en la Base de Datos**: Guarda los datos del formulario en una hoja de base de datos para su posterior seguimiento y análisis.
10. **Limpieza del Formulario**: Limpia el formulario para permitir nuevas entradas de datos.
11. **Confirmación de Envío**: Muestra un mensaje de confirmación al usuario indicando que el formulario ha sido enviado con éxito.

### Funciones Auxiliares

- `getSetupValue(sheet, cell)`: Obtiene el valor de una celda específica en la hoja de configuración.
- `getFormData(sheet)`: Recopila todos los datos de la hoja del formulario.
- `createHtmlTableWithStyles(sheet, data)`: Crea una tabla HTML estilizada con los datos del formulario.
- `sendEmail(to, cc, bcc, subject, body, attachments)`: Envía un correo electrónico con los destinatarios, asunto, cuerpo y adjuntos proporcionados.
- `saveToBase(sheet, data, newId)`: Guarda los datos del formulario en la hoja de base de datos con un nuevo identificador.
- `limpiar_form(sheet)`: Limpia los datos del formulario.
- `isValidEmail(email)`: Verifica si las direcciones de correo electrónico son válidas.
- `showConfirmationMessage()`: Muestra un mensaje de confirmación cuando el formulario ha sido enviado exitosamente.
- `showEmptyFormWarning()`: Muestra una advertencia cuando se intenta enviar un formulario vacío.
- `isFormEmpty(data)`: Verifica si el formulario está vacío.
- `getFileIdFromUrl(url)`: Obtiene el ID del archivo de la URL de Google Drive.

## Beneficios

- **Automatización**: Reduce la necesidad de intervención manual en el envío de correos y almacenamiento de datos.
- **Eficiencia**: Acelera el proceso de gestión de datos y garantiza que la información se procese de manera oportuna.
- **Validación**: Asegura que solo se envíen correos a direcciones válidas, minimizando errores.
- **Organización**: Mantiene una base de datos interna ordenada con todos los envíos de formularios, facilitando el seguimiento y análisis de la información.
- **Usabilidad**: Proporciona retroalimentación al usuario sobre el estado del formulario, mejorando la experiencia del usuario.

## Uso

1. **Configuración Inicial**: Asegúrate de tener una hoja de configuración ('Setup') con los campos necesarios para el correo electrónico (asunto, destinatarios, etc.). También, añade una celda para la URL del archivo adjunto si es necesario.
2. **Formulario**: Diseña tu hoja de formulario ('Form') donde los usuarios ingresarán los datos.
3. **Base de Datos**: Crea una hoja de base de datos ('Base') para almacenar los envíos del formulario.
4. **Ejecución del Script**: Ejecuta la función `enviar_form()` para procesar y enviar los datos del formulario.

Este flujo de trabajo proporciona una solución completa para la gestión de formularios en Google Sheets, combinando automatización, validación y almacenamiento eficiente de datos.
