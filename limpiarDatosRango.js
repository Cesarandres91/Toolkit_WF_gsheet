function limpiarDatosRango() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // Solicitar al usuario el rango de columnas
  var primeraCol = ui.prompt(
    'Rango de limpieza',
    'Ingrese la letra de la primera columna del rango:',
    ui.ButtonSet.OK_CANCEL
  );
  if (primeraCol.getSelectedButton() != ui.Button.OK) {
    return;
  }
  primeraCol = primeraCol.getResponseText().toUpperCase();

  var ultimaCol = ui.prompt(
    'Rango de limpieza',
    'Ingrese la letra de la última columna del rango:',
    ui.ButtonSet.OK_CANCEL
  );
  if (ultimaCol.getSelectedButton() != ui.Button.OK) {
    return;
  }
  ultimaCol = ultimaCol.getResponseText().toUpperCase();

  // Convertir letras de columna a números
  var primeraColNum = columnToNumber(primeraCol);
  var ultimaColNum = columnToNumber(ultimaCol);

  // Obtener el rango
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(1, primeraColNum, lastRow, ultimaColNum - primeraColNum + 1);

  // Obtener todos los valores del rango
  var values = range.getValues();

  // Limpiar los datos
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      var value = values[i][j];
      // Reemplazar null por cadena vacía y eliminar espacios en blanco
      if (value === null || value === undefined) {
        values[i][j] = '';
      } else if (typeof value === 'string') {
        values[i][j] = value.trim();
      }
    }
  }

  // Establecer los valores limpios de vuelta en el rango
  range.setValues(values);

  ui.alert('Limpieza de datos completada para el rango ' + primeraCol + ':' + ultimaCol);
}

// Función auxiliar para convertir letra de columna a número
function columnToNumber(column) {
  var result = 0;
  for (var i = 0; i < column.length; i++) {
    result *= 26;
    result += column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return result;
}
