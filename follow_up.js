function fetchEmailsAndUpdateSheet() {
  const LABEL_NAME = 'LABEL'; // Nombre de la etiqueta de Gmail a procesar
  const SHEET_ID = 'ID_HOJA'; // ID de la hoja de cálculo de Google Sheets
  const REVIEWED_LABEL_NAME = 'Revisado'; // Nombre de la etiqueta a añadir después de procesar
  const NO_READ_LABEL_NAME = 'NO_LEER'; // Nombre de la etiqueta para omitir correos
  const MAX_BODY_LENGTH = 5000; // Longitud máxima del cuerpo del mensaje truncado
  const BATCH_SIZE = 100; // Tamaño del lote para procesamiento
  const userEmail = Session.getActiveUser().getEmail(); // Email del usuario
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const startTime = new Date();
  const today = new Date();
  const lista_roja = "test@test.com, test2@test.com"; // Lista roja de correos separados por comas

  // Borrar toda la información de la hoja
  sheet.clear();

  // Agregar encabezados en la hoja de cálculo si no existen
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date_process', 'Tags', 'Email_ID', 'Subject', 'First_from', 'Last_From', 'To', 'CC', 'Date_first_email', 'Date_last_email', 'Body_First', 'Body_Last', 'Link', 'Total_days', 'Days_follow', 'Status_1', 'Status_2', 'Follow_up_template']);
  }

  try {
    const threads = GmailApp.getUserLabelByName(LABEL_NAME).getThreads();
    const reviewedLabel = GmailApp.getUserLabelByName(REVIEWED_LABEL_NAME);
    const noReadLabel = GmailApp.getUserLabelByName(NO_READ_LABEL_NAME);
    const existingEmailIds = getEmailIdsFromSheet(sheet);
    const listaRojaArray = lista_roja.split(',').map(email => email.trim().toLowerCase());
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
          const lastFrom = lastMessage.getFrom();
          const to = firstMessage.getTo();
          const cc = firstMessage.getCc();
          const firstMessageDate = firstMessage.getDate();
          const lastMessageDate = lastMessage.getDate();

          const lastMessageBody = cleanHtml(lastMessage.getBody(), MAX_BODY_LENGTH, true);
          const firstMessageBody = cleanHtml(firstMessage.getBody(), MAX_BODY_LENGTH, false);
          const processingDate = new Date();
          const senderName = extractSenderName(from);
          const followUpBody = `Hola ${senderName}, ¿Cómo estás?, ¿Pudiste revisarlo?`;
          const dias1 = calculateBusinessDays(firstMessageDate, today);
          const diasf = calculateBusinessDays(lastMessageDate, today);
          const defaultStatus = 'En_proceso';
          const tags = thread.getLabels().map(label => label.getName()).join(';');
          const link = getFirstMessageLink(thread);
          let status2;

          if (firstMessageDate.toDateString() === lastMessageDate.toDateString()) {
            status2 = 'No_respond';
          } else {
            status2 = 'On_track';
            const lastFromEmail = extractEmail(lastFrom).toLowerCase();
            if (listaRojaArray.includes(lastFromEmail)) {
              status2 = 'Follow_up';
            }
          }

          const row = [processingDate, tags, emailId, subject, from, lastFrom, to, cc, firstMessageDate, lastMessageDate, firstMessageBody, lastMessageBody, link, dias1, diasf, defaultStatus, status2, followUpBody];

          sheet.appendRow(row);
          processedCount++;

          // Añadir la etiqueta "Revisado" al hilo
          thread.addLabel(reviewedLabel);
        } catch (error) {
          Logger.log('Error procesando el hilo: ' + error.message);
          // Agrega el error en la hoja para facilitar la depuración
          sheet.appendRow(['ERROR', '', '', '', '', '', '', '', '', '', '', new Date(), '', `Error procesando el hilo: ${error.message}`, '', '', '', '']);
        }
      });
      Utilities.sleep(2000);  // Pausa para evitar exceder los límites de tiempo de ejecución
    }

    const endTime = new Date();
    const duration = (endTime - startTime) / 1000; // Duración en segundos
    sendCompletionEmail(userEmail, processedCount, duration, SHEET_ID);

    // Forzar el tamaño de todas las filas a 21 y establecer el ajuste de texto en recorte
    const lastRow = sheet.getLastRow();
    setRowHeightsForcedAndClip(sheet, 1, lastRow, 21);

  } catch (error) {
    Logger.log('Error en la función principal: ' + error.message);
    // Agrega el error en la hoja para facilitar la depuración
    sheet.appendRow(['ERROR', '', '', '', '', '', '', '', '', '', '', new Date(), '', `Error en la función principal: ${error.message}`, '', '', '', '']);
  }
}

function cleanHtml(html, maxLength, prioritizeEnd) {
  // Crear un documento HTML para eliminar las etiquetas HTML
  const plainText = HtmlService.createHtmlOutput(html).getContent();
  // Eliminar etiquetas de estilo y script
  let plainTextNoStyleScript = plainText.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '').replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
  // Convertir el HTML a texto plano
  let plainTextConverted = plainTextNoStyleScript.replace(/<\/?[^>]+(>|$)/g, "");

  // Reducir saltos de línea a uno solo
  plainTextConverted = plainTextConverted.replace(/(\r\n|\n|\r){2,}/g, '\n');

  // Reemplazar espacios en blanco mayores a 5 por un salto de línea
  plainTextConverted = plainTextConverted.replace(/ {5,}/g, '\n');

  // Reducir espacios en blanco intermedios mayores a 3 a dos espacios
  plainTextConverted = plainTextConverted.replace(/ {3,4}/g, '  ');

  // Eliminar líneas vacías adicionales y reducir saltos de línea excesivos
  plainTextConverted = plainTextConverted.replace(/\n{2,}/g, '\n\n');

  // Truncar el texto
  if (plainTextConverted.length > maxLength) {
    if (prioritizeEnd) {
      return plainTextConverted.slice(-maxLength);
    } else {
      return plainTextConverted.slice(0, maxLength);
    }
  }
  return plainTextConverted;
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

function extractEmail(fromField) {
  const emailPattern = /<(.+?)>/;
  const match = emailPattern.exec(fromField);
  if (match) {
    return match[1];
  } else {
    return fromField; // Devuelve el campo completo si no coincide el patrón
  }
}

function getFirstMessageLink(thread) {
  const firstMessage = thread.getMessages()[0];
  return firstMessage.getId() ? `https://mail.google.com/mail/u/0/#inbox/${firstMessage.getId()}` : '';
}

function sendCompletionEmail(userEmail, processedCount, duration, sheetId) {
  const subject = 'Gmail Processing Completed';
  const body = `Hola,\n\nEl procesamiento de correos ha finalizado.\n\nTotal de correos procesados: ${processedCount}\nDuración: ${duration} segundos\n\nPuedes ver la planilla en el siguiente enlace: https://docs.google.com/spreadsheets/d/${sheetId}\n\nSaludos,\nTu Script de Google Apps`;
  MailApp.sendEmail(userEmail, subject, body);
}

function setRowHeightsForcedAndClip(sheet, startRow, numRows, height) {
  for (let i = startRow; i <= numRows; i++) {
    sheet.setRowHeight(i, height);
  }
  const range = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

function calculateBusinessDays(startDate, endDate) {
  let count = 0;
  let currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    const dayOfWeek = currentDate.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Excluir domingos (0) y sábados (6)
      count++;
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }
  return count;
}
