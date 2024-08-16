function procesarDatos() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const hojaBase = SS.getSheetByName("base");
  const LIBRO_ORIGEN_ID = "MI_ID"; // Reemplaza con el ID real
  const BATCH_SIZE = 1000;
  const COLUMNA_FILTRO = "A";
  const INDICE_COLUMNA_FILTRO = 0;

  if (!hojaBase) {
    Logger.log("Error: La hoja 'base' no existe.");
    return;
  }

  // Obtener la lista de valores
  const ultimaFila = hojaBase.getLastRow();
  const lista = hojaBase.getRange(`${COLUMNA_FILTRO}2:${COLUMNA_FILTRO}${ultimaFila}`).getValues()
    .flat()
    .filter(String)
    .map(value => value.toString().trim());
  const listaSet = new Set(lista);

  let libroOrigen;
  try {
    libroOrigen = SpreadsheetApp.openById(LIBRO_ORIGEN_ID);
  } catch (error) {
    Logger.log("Error al abrir el libro origen: " + error.message);
    return;
  }

  const hojaDatosOrigen = libroOrigen.getSheetByName("datos_origen");
  if (!hojaDatosOrigen) {
    Logger.log("Error: La hoja 'datos_origen' no existe en el libro origen.");
    return;
  }

  // Preparar la hoja de destino
  let hojaDatos01 = SS.getSheetByName("datos_01");
  if (!hojaDatos01) {
    hojaDatos01 = SS.insertSheet("datos_01");
  } else {
    hojaDatos01.clear();
  }

  // Crear o limpiar la hoja "Errores"
  let hojaErrores = SS.getSheetByName("Errores");
  if (!hojaErrores) {
    hojaErrores = SS.insertSheet("Errores");
  } else {
    hojaErrores.clearContents();
  }

  // Obtener todos los datos de origen
  const datosOrigen = hojaDatosOrigen.getDataRange().getValues();
  const numColumnas = datosOrigen[0].length;

  let datosFiltrados = [];
  let esEncabezado = true;
  const errores = [];

  // Filtrar datos y registrar errores
  datosOrigen.forEach((fila, index) => {
    const valorFiltroOrigen = fila[INDICE_COLUMNA_FILTRO].toString().trim();

    if (esEncabezado) {
      datosFiltrados.push(fila);
      esEncabezado = false;
    } else if (listaSet.has(valorFiltroOrigen)) {
      datosFiltrados.push(fila);
    } else {
      errores.push([valorFiltroOrigen, "No encontrado"]);
    }

    // Escribir datos filtrados en lotes
    if (datosFiltrados.length >= BATCH_SIZE || index === datosOrigen.length - 1) {
      if (datosFiltrados.length > 0) {
        hojaDatos01.getRange(hojaDatos01.getLastRow() + 1, 1, datosFiltrados.length, numColumnas).setValues(datosFiltrados);
      }
      datosFiltrados = [];
    }

    // Pausa para evitar timeouts en ejecuciones largas
    if (index % (BATCH_SIZE * 10) === 0) {
      Utilities.sleep(1000);
      SpreadsheetApp.flush();
    }
  });

  // Escribir los errores al final
  if (errores.length > 0) {
    hojaErrores.getRange(1, 1, errores.length, errores[0].length).setValues(errores);
  }

  SpreadsheetApp.flush();
  Logger.log("Proceso completado");
}

