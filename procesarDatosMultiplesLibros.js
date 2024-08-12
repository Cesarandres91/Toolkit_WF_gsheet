function procesarDatosMultiplesLibros() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const hojaBase = SS.getSheetByName("base");
  const BATCH_SIZE = 1000; // Tamaño del lote para procesar y escribir datos

  // IDs de los libros origen (reemplaza con los IDs reales)
  const LIBROS_ORIGEN_IDS = [
    "ID_LIBRO_1",
    "ID_LIBRO_2",
    "ID_LIBRO_3",
    "ID_LIBRO_4"
  ];

  // Obtener la lista de valores
  const ultimaFila = hojaBase.getLastRow();
  const lista = hojaBase.getRange("A2:A" + ultimaFila).getValues().flat();
  const listaSet = new Set(lista); // Convertir a Set para búsqueda más rápida

  // Procesar cada libro
  LIBROS_ORIGEN_IDS.forEach((libroID, index) => {
    procesarLibro(libroID, listaSet, index + 1);
  });

  Logger.log("Proceso completado para todos los libros");
}

function procesarLibro(libroID, listaSet, numeroLibro) {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const libroOrigen = SpreadsheetApp.openById(libroID);
  const hojaDatosOrigen = libroOrigen.getSheetByName("datos_origen");

  // Preparar la hoja de destino
  const nombreHojaDestino = `datos_0${numeroLibro}`;
  let hojaDestino = SS.getSheetByName(nombreHojaDestino);
  if (!hojaDestino) {
    hojaDestino = SS.insertSheet(nombreHojaDestino);
  } else {
    hojaDestino.clear();
  }

  // Obtener datos de origen en lotes y filtrar
  const numColumnas = hojaDatosOrigen.getLastColumn();
  const numFilas = hojaDatosOrigen.getLastRow();
  let datosFiltrados = [];
  let filaInicio = 1;

  while (filaInicio <= numFilas) {
    const filasALeer = Math.min(BATCH_SIZE, numFilas - filaInicio + 1);
    const rangoLectura = hojaDatosOrigen.getRange(filaInicio, 1, filasALeer, numColumnas);
    const datos = rangoLectura.getValues();

    const datosFiltradosLote = datos.filter((fila, index) => {
      return filaInicio === 1 || listaSet.has(fila[0]);
    });

    datosFiltrados = datosFiltrados.concat(datosFiltradosLote);
    filaInicio += BATCH_SIZE;

    // Escribir datos filtrados en lotes
    if (datosFiltrados.length >= BATCH_SIZE || filaInicio > numFilas) {
      hojaDestino.getRange(hojaDestino.getLastRow() + 1, 1, datosFiltrados.length, numColumnas).setValues(datosFiltrados);
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
    hojaDestino.getRange(hojaDestino.getLastRow() + 1, 1, datosFiltrados.length, numColumnas).setValues(datosFiltrados);
  }

  SpreadsheetApp.flush();
  Logger.log(`Proceso completado para el libro ${numeroLibro}`);
}
