function procesarDatos() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const hojaBase = SS.getSheetByName("base");
  const LIBRO_ORIGEN_ID = "MI_ID"; // Reemplaza "MI_ID" con el ID real del libro origen
  const BATCH_SIZE = 1000; // Tamaño del lote para procesar y escribir datos

  // Obtener la lista de valores
  const ultimaFila = hojaBase.getLastRow();
  const lista = hojaBase.getRange("A2:A" + ultimaFila).getValues().flat().filter(String);
  const listaSet = new Set(lista); // Convertir a Set para búsqueda más rápida

  // Abrir el libro origen
  const libroOrigen = SpreadsheetApp.openById(LIBRO_ORIGEN_ID);
  const hojaDatosOrigen = libroOrigen.getSheetByName("datos_origen");

  // Preparar la hoja de destino
  let hojaDatos01 = SS.getSheetByName("datos_01");
  if (!hojaDatos01) {
    hojaDatos01 = SS.insertSheet("datos_01");
  } else {
    hojaDatos01.clear();
  }

  // Obtener datos de origen en lotes y filtrar
  const numColumnas = hojaDatosOrigen.getLastColumn();
  const numFilas = hojaDatosOrigen.getLastRow();
  let datosFiltrados = [];
  let filaInicio = 1;
  let esEncabezado = true;

  while (filaInicio <= numFilas) {
    const filasALeer = Math.min(BATCH_SIZE, numFilas - filaInicio + 1);
    const rangoLectura = hojaDatosOrigen.getRange(filaInicio, 1, filasALeer, numColumnas);
    const datos = rangoLectura.getValues();

    const datosFiltradosLote = datos.filter((fila, index) => {
      if (esEncabezado) {
        esEncabezado = false;
        return true; // Siempre incluir la primera fila (encabezados)
      }
      return listaSet.has(fila[0]) && fila[0] !== "";
    });

    datosFiltrados = datosFiltrados.concat(datosFiltradosLote);
    filaInicio += BATCH_SIZE;

    // Escribir datos filtrados en lotes
    if (datosFiltrados.length >= BATCH_SIZE || filaInicio > numFilas) {
      if (datosFiltrados.length > 0) {
        hojaDatos01.getRange(hojaDatos01.getLastRow() + 1, 1, datosFiltrados.length, numColumnas).setValues(datosFiltrados);
      }
      datosFiltrados = [];
    }

    // Pausa para evitar timeouts en ejecuciones largas
    if (filaInicio % (BATCH_SIZE * 10) === 0) {
      Utilities.sleep(1000);
      SpreadsheetApp.flush();
    }
  }

  // Escribir cualquier dato restante
  if (datosFiltrados.length > 0) {
    hojaDatos01.getRange(hojaDatos01.getLastRow() + 1, 1, datosFiltrados.length, numColumnas).setValues(datosFiltrados);
  }

  SpreadsheetApp.flush();
  Logger.log("Proceso completado");
}
