var hhtData = "HHT";

function obtenerMatrizHhtPaginado(offset, limit, filtro = "") {
  const hoja = getSpreadsheetPersonal().getSheetByName(hhtData);
  const lastRow = hoja.getLastRow();

  if (lastRow < 2) {
    return { headersHht: [], filas: [], total: 0 };
  }

  // Leer encabezados desde la primera fila
  const headersHht = hoja.getRange(1, 1, 1, 9).getDisplayValues()[0];

  // Leer solo datos desde la fila 2 hasta la última (sin encabezado)
  const numFilas = lastRow - 1;
  const datos = hoja.getRange(2, 1, numFilas, 9).getDisplayValues();

  // Invertimos los datos para mostrar los últimos primero
  let filas = datos.reverse();

  // Filtro si se proporciona
  if (filtro) {
    const texto = filtro.toLowerCase();
    filas = filas.filter(fila =>
      fila.some(celda => String(celda).toLowerCase().includes(texto))
    );
  }

  // Paginación
  const paginadas = filas.slice(offset, offset + limit);

  return {
    headersHht,
    filas: paginadas,
    total: filas.length
  };
}



function agregarPreguntaHht(data) {
  const hoja = getSpreadsheetPersonal().getSheetByName(hhtData);
  
  // Generar ID único basado en timestamp (7 dígitos + prefijo "E")
  const timestamp = Date.now().toString().slice(-7);
  const idUnico = "E" + timestamp;

  data[0] = idUnico; // Columna A
  hoja.appendRow(data);
}


function actualizarHht(data) {
  const hoja = getSpreadsheetPersonal().getSheetByName(hhtData);
  const id = String(data[0]).trim();
  const lastRow = hoja.getLastRow();

  if (lastRow < 2) return false;

  // Leer solo columna A (ID) desde fila 2 hacia abajo
  const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) {
      hoja.getRange(i + 2, 1, 1, data.length).setValues([data]); // +2 para ajustar fila real
      return true;
    }
  }
  return false;
}

function eliminarHhtPorID(id) {
  const hoja = getSpreadsheetPersonal().getSheetByName(hhtData);
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return false;

  // Leer solo la columna A desde la fila 2
  const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();

  const idBuscado = String(id).trim();

  for (let i = 0; i < ids.length; i++) {
    const idFila = String(ids[i][0]).trim();
    if (idFila === idBuscado) {
      hoja.deleteRow(i + 2); // +2 porque comenzamos en fila 2
      return true;
    }
  }
  return false;
}
function listarHojas() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(s => Logger.log(s.getName()));
}
