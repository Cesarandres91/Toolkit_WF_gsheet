function enviar_form() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName('Setup');
  var formSheet = ss.getSheetByName('Form');
  var baseSheet = ss.getSheetByName('Base');

  // Obtener los datos de la hoja Setup
  var subject = getSetupValue(setupSheet, 'B2');
  var to = getSetupValue(setupSheet, 'B3');
  var cc = getSetupValue(setupSheet, 'B4');
  var bcc = getSetupValue(setupSheet, 'B5');
  var body1 = getSetupValue(setupSheet, 'B6');
  var addIdToSubject = getSetupValue(setupSheet, 'B7');
  var currentId = getSetupValue(setupSheet, 'B9');
  var attachmentUrl = getSetupValue(setupSheet, 'B10'); // URL del archivo adjunto

  // Verificar correos electrónicos válidos
  if (!isValidEmail(to) || !isValidEmail(cc) || !isValidEmail(bcc)) {
    Logger.log('Invalid email address found');
    return;
  }

  // Obtener la tabla de la hoja Form
  var formData = getFormData(formSheet);

  // Verificar si el formulario está vacío
  if (isFormEmpty(formData)) {
    showEmptyFormWarning();
    return; // Salir de la función si el formulario está vacío
  }

  // Obtener el nuevo Id y actualizarlo en Setup
  var newId = parseInt(currentId, 10);
  setupSheet.getRange('B9').setValue(newId + 1);

  // Añadir el Id al subject si corresponde
  if (addIdToSubject.toUpperCase() === "SI" || addIdToSubject.toUpperCase() === "OK") {
    subject = subject + " " + newId;
  }

  // Crear el cuerpo del correo
  var body = body1 + '<br>' + createHtmlTableWithStyles(formSheet, formData);

  // Preparar el archivo adjunto si está disponible
  var attachments = [];
  if (attachmentUrl) {
    try {
      var file = DriveApp.getFileById(getFileIdFromUrl(attachmentUrl));
      attachments.push(file.getAs(MimeType.PDF));
    } catch (e) {
      Logger.log('Error fetching attachment: ' + e.message);
    }
  }

  // Enviar el correo
  sendEmail(to, cc, bcc, subject, body, attachments);

  // Guardar la información en la hoja Base
  saveToBase(baseSheet, formData, newId);

  // Limpiar la hoja Form
  limpiar_form(formSheet);

  // Confirmación de envío
  showConfirmationMessage();
}

function getSetupValue(sheet, cell) {
  return sheet.getRange(cell).getValue();
}

function getFormData(sheet) {
  var range = sheet.getDataRange();
  return range.getValues();
}

function createHtmlTableWithStyles(sheet, data) {
  var htmlTable = '<table border="1" style="border-collapse:collapse">';
  var firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var styles = getRowStyles(firstRow);

  // Añadir los estilos de la primera fila (títulos)
  htmlTable += '<tr>';
  for (var j = 0; j < data[0].length; j++) {
    htmlTable += '<th style="' + styles[j] + '">' + data[0][j] + '</th>';
  }
  htmlTable += '</tr>';

  // Añadir los datos restantes
  for (var i = 1; i < data.length; i++) {
    htmlTable += '<tr>';
    for (var j = 0; j < data[i].length; j++) {
      htmlTable += '<td>' + data[i][j] + '</td>';
    }
    htmlTable += '</tr>';
  }
  htmlTable += '</table>';
  return htmlTable;
}

function getRowStyles(rowRange) {
  var styles = [];
  var backgrounds = rowRange.getBackgrounds()[0];
  var fonts = rowRange.getFontFamilies()[0];
  var fontSizes = rowRange.getFontSizes()[0];
  var fontWeights = rowRange.getFontWeights()[0];
  var fontColors = rowRange.getFontColors()[0];

  for (var i = 0; i < backgrounds.length; i++) {
    var style = 'background-color:' + backgrounds[i] + ';';
    style += 'font-family:' + fonts[i] + ';';
    style += 'font-size:' + fontSizes[i] + 'px;';
    style += 'font-weight:' + fontWeights[i] + ';';
    style += 'color:' + fontColors[i] + ';';
    styles.push(style);
  }
  return styles;
}

function sendEmail(to, cc, bcc, subject, body, attachments) {
  var mailOptions = {
    to: to,
    cc: cc,
    bcc: bcc,
    subject: subject,
    htmlBody: body
  };
  if (attachments && attachments.length > 0) {
    mailOptions.attachments = attachments;
  }
  MailApp.sendEmail(mailOptions);
}

function saveToBase(sheet, data, newId) {
  var currentDate = new Date();
  var userEmail = Session.getActiveUser().getEmail();

  // Inserta filas para desplazar los datos existentes hacia abajo
  sheet.insertRowsBefore(2, data.length - 1);

  var rows = [];
  for (var i = 1; i < data.length; i++) {  // Start from 1 to skip header
    var row = [newId, i, userEmail, currentDate];  // Add newId and Id_R
    for (var j = 0; j < data[i].length; j++) {
      row.push(data[i][j]);
    }
    row.push('Pendiente');
    rows.push(row);
  }

  // Rango para las nuevas filas que se van a insertar
  var newRange = sheet.getRange(2, 1, rows.length, rows[0].length);

  // Eliminar las validaciones de datos existentes en el rango
  newRange.clearDataValidations();

  // Establecer los valores en la hoja
  newRange.setValues(rows);

  // Añadir validación de datos para la columna de estado
  var lastRow = sheet.getLastRow();
  var numColumns = data[0].length;
  var range = sheet.getRange(2, numColumns + 5, lastRow - 1, 1);  // Adjusted column index for validation
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pendiente', 'OK', 'Comentario', 'Rechazado'])
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function limpiar_form(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow > 1) {
    var range = sheet.getRange('A2:' + sheet.getRange(1, lastColumn).getA1Notation()[0] + lastRow);
    range.clearContent();
  }
}

function isValidEmail(email) {
  var emailPattern = /^[^\s@]+@[^\s@]+$/;
  if (email) {
    var emails = email.split(',').map(function(e) { return e.trim(); });
    for (var i = 0; i < emails.length; i++) {
      if (!emailPattern.test(emails[i])) {
        return false;
      }
    }
  }
  return true;
}

//Aviso de confirmación que se ha enviado el formulario
function showConfirmationMessage() {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutput('<div style="text-align:center;">' +
    '<h2 style="color:green;">Formulario enviado exitosamente</h2>' +
    '<p style="font-size:16px;">Tu formulario ha sido enviado y la base de datos se ha actualizado correctamente.</p>' +
    '<p style="font-size:16px;">Muchas gracias por tu colaboración.</p>' +
    '<div style="display: inline-block; border-radius: 50%; width: 70px; height: 70px; background-color: green; position: relative;">' +
    '  <div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); width: 25px; height: 15px; border-left: 5px solid white; border-bottom: 5px solid white; transform: translate(-50%, -50%) rotate(-45deg);"></div>' +
    '</div>' +
    '</div>')
    .setWidth(400)
    .setHeight(300);
  ui.showModalDialog(htmlOutput, 'Confirmación');
}

//Aviso de no enviar el formulario vacío
function showEmptyFormWarning() {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutput('<div style="text-align:center;">' +
    '<h2 style="color:red;">No se puede enviar el formulario</h2>' +
    '<p style="font-size:16px;">El formulario está vacío. Por favor, completa la información requerida antes de enviarlo.</p>' +
    '<div style="display: inline-block; border-radius: 50%; width: 70px; height: 70px; background-color: red; position: relative;">' +
    '  <div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); width: 20px; height: 20px; border-left: 5px solid white; border-top: 5px solid white; transform: translate(-50%, -50%) rotate(45deg);"></div>' +
    '</div>' +
    '</div>')
    .setWidth(400)
    .setHeight(300);
  ui.showModalDialog(htmlOutput, 'Advertencia');
}

//Revisar si el formulario esta vacío
function isFormEmpty(data) {
  for (var i = 1; i < data.length; i++) { // Empezar en 1 para saltar la fila de títulos
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] !== '') {
        return false;
      }
    }
  }
  return true;
}

// Obtener el ID del archivo de la URL de Google Drive
function getFileIdFromUrl(url) {
  var fileIdMatch = url.match(/[-\w]{25,}/);
  return fileIdMatch ? fileIdMatch[0] : null;
}
