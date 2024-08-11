function rellenarVacios() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getActiveRange();
  
  // Comprobar si se ha seleccionado algo
  if (!range) {
    SpreadsheetApp.getUi().alert('Por favor, selecciona una columna o rango.');
    return;
  }
  
  // Obtener las dimensiones del rango seleccionado
  var numColumns = range.getNumColumns();
  var numRows = range.getNumRows();
  var startRow = range.getRow();
  var startCol = range.getColumn();
  
  // Iterar por cada columna en el rango seleccionado
  for (var col = 0; col < numColumns; col++) {
    var lastValue = sheet.getRange(startRow, startCol + col).getValue();
    var lastRow = sheet.getLastRow();
    
    // Iterar por cada celda en la columna
    for (var row = startRow; row <= lastRow; row++) {
      var cell = sheet.getRange(row, startCol + col);
      var value = cell.getValue();
      
      if (value === '') {
        cell.setValue(lastValue);
      } else {
        lastValue = value;
      }
    }
  }
  
  SpreadsheetApp.getUi().alert('Proceso completado.');
}
