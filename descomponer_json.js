function desempaquetarJSON() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base");
  var data = sheet.getDataRange().getValues();
  var headers = [];
  var outputData = [];

  for (var i = 0; i < data.length; i++) {
    var jsonString = data[i][0];
    try {
      var jsonObject = JSON.parse(jsonString);
      var flattenedObject = flattenObject(jsonObject);
      
      if (i === 0) {
        headers = Object.keys(flattenedObject);
        outputData.push(headers);
      }
      
      var row = headers.map(function(header) {
        return flattenedObject[header] || "";
      });
      outputData.push(row);
    } catch (e) {
      Logger.log("Error en la fila " + (i+1) + ": " + e.message);
    }
  }
  
  // Crear una nueva hoja para los resultados
  var outputSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Resultados_Desempaquetados");
  outputSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
}

function flattenObject(obj, prefix = '') {
  return Object.keys(obj).reduce((acc, k) => {
    const pre = prefix.length ? prefix + '_' : '';
    if (typeof obj[k] === 'object' && obj[k] !== null && !Array.isArray(obj[k])) {
      Object.assign(acc, flattenObject(obj[k], pre + k));
    } else if (Array.isArray(obj[k])) {
      obj[k].forEach((item, index) => {
        if (typeof item === 'object' && item !== null) {
          Object.assign(acc, flattenObject(item, `${pre}${k}_${index}`));
        } else {
          acc[`${pre}${k}_${index}`] = item;
        }
      });
    } else {
      acc[pre + k] = obj[k];
    }
    return acc;
  }, {});
}
