function obtenerAvisosERP() {
  const sheet = getSpreadsheetPersonal().getSheetByName("AVISOS");
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return []; // No hay datos

  // Ahora leemos 5 columnas: B2:F (Categoría, Título, Descripción, URL, Fecha)
  const data = sheet.getRange(2, 2, lastRow - 1, 5).getValues(); // B2:F

  return data.map((row, index) => ({
    id: index + 1, // ID único para cada registro
    categoria: row[0], // Columna B: 0 = categoría principal, >0 = subcategoría
    titulo: row[1],    // Columna C
    descripcion: row[2], // Columna D
    url: row[3],       // Columna E
    fecha: row[4]
      ? Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), "dd/MM/yyyy")
      : "",            // Columna F
    esCategoria: row[0] === 0 || row[0] === "0" // Identificar si es categoría principal
  }));
}
