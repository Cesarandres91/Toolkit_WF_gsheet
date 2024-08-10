function prosesajsonversion2() {
  const BATCH_SIZE = 50000; //más es rapido pero menos reduce la memria
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = new Set();
  const output = [];

  function flattenJSON(obj, prefix = '') {
    const result = {};
    for (const [key, value] of Object.entries(obj)) {
      const newKey = prefix ? `${prefix}.${key}` : key;
      if (typeof value === 'object' && value !== null) {
        Object.assign(result, flattenJSON(value, newKey));
      } else {
        result[newKey] = value;
      }
    }
    return result;
  }

  function processBatch(startRow, endRow) {
    for (let i = startRow; i < endRow && i < data.length; i++) {
      const jsonString = data[i][0];
      try {
        const jsonObj = JSON.parse(jsonString);
        const flatObj = flattenJSON(jsonObj);
        
        Object.keys(flatObj).forEach(key => headers.add(key));
        output.push(flatObj);
      } catch (e) {
        Logger.log(`Error en la fila ${i + 1}: ${e.message}`);
      }
    }
  }

  // Procesar datos en lotes
  for (let i = 0; i < data.length; i += BATCH_SIZE) {
    processBatch(i, i + BATCH_SIZE);
    
    // Pausa para evitar timeout
    if (i % (BATCH_SIZE * 10) === 0) {
      SpreadsheetApp.flush();
      Utilities.sleep(1000);
    }
  }

  const headerArray = Array.from(headers);
  const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  newSheet.getRange(1, 1, 1, headerArray.length).setValues([headerArray]);

  // Escribir datos en lotes
  for (let i = 0; i < output.length; i += BATCH_SIZE) {
    const batch = output.slice(i, i + BATCH_SIZE).map(obj => 
      headerArray.map(header => obj[header] !== undefined ? obj[header] : '')
    );
    newSheet.getRange(i + 2, 1, batch.length, headerArray.length).setValues(batch);
    
    // Pausa para evitar timeout
    if (i % (BATCH_SIZE * 10) === 0) {
      SpreadsheetApp.flush();
      Utilities.sleep(1000);
    }
  }
}
