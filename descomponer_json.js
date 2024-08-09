function parseAndWriteJSON() {
  var json = {
    "clave": "hola",
    "detalles": {
      "nombre": "Juan",
      "edad": 30,
      "hobbies": ["futbol", "lectura"]
    },
    "ubicaciones": [
      {"ciudad": "Santiago", "pais": "Chile"},
      {"ciudad": "Buenos Aires", "pais": "Argentina"}
    ]
  };
  
  // Obtener la hoja activa
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Llamar a la función para descomponer el JSON y escribir en la hoja
  writeJsonToSheet(sheet, json);
}

function writeJsonToSheet(sheet, json, prefix = '') {
  var row = 1;
  var col = 1;
  
  for (var key in json) {
    if (typeof json[key] === 'object' && !Array.isArray(json[key])) {
      // Si es un objeto, llamar a la función recursivamente
      writeJsonToSheet(sheet, json[key], prefix + key + '_');
    } else if (Array.isArray(json[key])) {
      // Si es un arreglo, explotar los valores en columnas
      for (var i = 0; i < json[key].length; i++) {
        if (typeof json[key][i] === 'object') {
          writeJsonToSheet(sheet, json[key][i], prefix + key + i + '_');
        } else {
          sheet.getRange(row, col).setValue(prefix + key + i);
          sheet.getRange(row + 1, col).setValue(json[key][i]);
          col++;
        }
      }
    } else {
      // Si es un valor simple, escribir en la hoja
      sheet.getRange(row, col).setValue(prefix + key);
      sheet.getRange(row + 1, col).setValue(json[key]);
      col++;
    }
  }
}
