function fetchEmailsAndUpdateSheet() {
  const LABEL_NAME = 'LABEL_1'; // Nombre de la etiqueta de Gmail a procesar
  const SHEET_ID = 'MY_ID'; // ID de la hoja de cálculo de Google Sheets
  const REVIEWED_LABEL_NAME = 'Revisado'; // Nombre de la etiqueta a añadir después de procesar
  const NO_READ_LABEL_NAME = 'NO_LEER'; // Nombre de la etiqueta para omitir correos
  const MAX_BODY_LENGTH = 5000; // Longitud máxima del cuerpo del mensaje truncado
  const BATCH_SIZE = 100; // Tamaño del lote para procesamiento
  const userEmail = Session.getActiveUser().getEmail(); // Email del usuario
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const startTime = new Date();

  // Agregar encabezados en la hoja de cálculo si no existen
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Email ID', 'Asunto', 'De', 'Para', 'CC', 'Fecha Envio Primer Correo', 'Fecha Respuesta Último Correo', 'Snippet', 'Body Último Correo (Truncado)', 'Body Primer Correo (Truncado)', 'Estado', 'Fecha de Procesamiento', 'Propuesta de Follow-up', 'Errores']);
  }

  try {
    const threads = GmailApp.getUserLabelByName(LABEL_NAME).getThreads();
    const reviewedLabel = GmailApp.getUserLabelByName(REVIEWED_LABEL_NAME);
    const noReadLabel = GmailApp.getUserLabelByName(NO_READ_LABEL_NAME);
    const existingEmailIds = getEmailIdsFromSheet(sheet);
    let processedCount = 0;

    for (let i = 0; i < threads.length; i += BATCH_SIZE) {
      const batch = threads.slice(i, i + BATCH_SIZE);

      batch.forEach(thread => {
        try {
          // Omitir hilos con la etiqueta "NO_LEER"
          if (threadHasLabel(thread, noReadLabel)) {
            return;
          }

          const messages = thread.getMessages();
          const firstMessage = messages[0];
          const lastMessage = messages[messages.length - 1];
          const emailId = thread.getId();

          // Evitar procesar correos duplicados
          if (existingEmailIds.includes(emailId)) return;

          const subject = thread.getFirstMessageSubject();
          const from = firstMessage.getFrom();
          const to = firstMessage.getTo();
          const cc = firstMessage.getCc();
          const firstMessageDate = firstMessage.getDate();
          const lastMessageDate = lastMessage.getDate();
          const snippet = thread.getSnippet();
          const lastMessageBody = truncateHtmlBody(lastMessage.getBody(), MAX_BODY_LENGTH, true);
          const firstMessageBody = truncateHtmlBody(firstMessage.getBody(), MAX_BODY_LENGTH, false);
          const status = (firstMessage.getFrom() === userEmail) ? 'send_for_me' : 'other';
          const processingDate = new Date();
          const senderName = extractSenderName(from);
          const followUpBody = `Hola ${senderName}, ¿Cómo estás?, ¿Pudiste revisarlo?`;

          const row = [emailId, subject, from, to, cc, firstMessageDate, lastMessageDate, snippet, lastMessageBody, firstMessageBody, status, processingDate, followUpBody, ''];

          const lastRow = sheet.appendRow(row).getRow();
          processedCount++;
          
          // Añadir color según el estado
          if (status === 'send_for_me') {
            sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).setBackground('lightgreen');
          } else {
            sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).setBackground('lightcoral');
          }

          // Añadir la etiqueta "Revisado" al hilo
          thread.addLabel(reviewedLabel);
        } catch (error) {
          Logger.log('Error procesando el hilo: ' + error.message);
          // Agrega el error en la hoja para facilitar la depuración
          sheet.appendRow(['ERROR', '', '', '', '', '', '', '', '', '', '', new Date(), '', `Error procesando el hilo: ${error.message}`]);
        }
      });
      Utilities.sleep(2000);  // Pausa para evitar exceder los límites de tiempo de ejecución
    }

    const endTime = new Date();
    const duration = (endTime - startTime) / 1000; // Duración en segundos
    sendCompletionEmail(userEmail, processedCount, duration);

  } catch (error) {
    Logger.log('Error en la función principal: ' + error.message);
    // Agrega el error en la hoja para facilitar la depuración
    sheet.appendRow(['ERROR', '', '', '', '', '', '', '', '', '', '', new Date(), '', `Error en la función principal: ${error.message}`]);
  }
}

function truncateHtmlBody(htmlBody, maxLength, prioritizeEnd) {
  // Eliminar etiquetas HTML
  const plainText = htmlBody.replace(/<\/?[^>]+(>|$)/g, "");
  // Truncar el texto
  if (plainText.length > maxLength) {
    if (prioritizeEnd) {
      return plainText.slice(-maxLength);
    } else {
      return plainText.slice(0, maxLength);
    }
  }
  return plainText;
}

function getEmailIdsFromSheet(sheet) {
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => row[0]);  // Asumiendo que la columna 0 contiene los IDs de los correos
}

function threadHasLabel(thread, label) {
  return thread.getLabels().some(l => l.getName() === label.getName());
}

function extractSenderName(fromField) {
  const nameEmailPattern = /^(.*?)\s*<.*?>$/;
  const match = nameEmailPattern.exec(fromField);
  if (match) {
    return match[1];
  } else {
    return fromField; // Devuelve el campo completo si no coincide el patrón
  }
}

function sendCompletionEmail(userEmail, processedCount, duration) {
  const subject = 'Gmail Processing Completed';
  const body = `Hola,\n\nEl procesamiento de correos ha finalizado.\n\nTotal de correos procesados: ${processedCount}\nDuración: ${duration} segundos\n\nSaludos,\nTu Script de Google Apps`;
  MailApp.sendEmail(userEmail, subject, body);
}
