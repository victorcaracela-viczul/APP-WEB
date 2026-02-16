/* DEFINE GLOBAL VARIABLES, CHANGE THESE VARIABLES TO MATCH WITH YOUR SHEET */

const folderimgcheck = "13qGGx2VJRlbcPSw9b2Ldn10HlJQJg0zd" // ARCHIVO 2 "Imagen inspecciones"
let folderpdfcheck = "1Be7s5TlJRS6sj0f6NxhGPK9EqFJcVM6N" // ARCHIVO 3 "PDF inspecciones borrar"

let cachedCheck = null;
function getCheckSpreadsheet() {
  if (!cachedCheck) {
    cachedCheck = SpreadsheetApp.openById(SPREADSHEET_IDS.check);
  }
  return cachedCheck;
}

function getDatosRegistroCheck(offset, limit, search1, search2, columnaFiltro1, columnaFiltro2) {
  try {
    const hoja = getCheckSpreadsheet().getSheetByName("B DATOS");
    const lastRow = hoja.getLastRow();
    // ✅ Leer hasta columna CT (98) para incluir fotos de observación y levantamiento
    const datos = hoja.getRange(1, 1, lastRow, 98).getDisplayValues();

    const headers = datos[0].slice(0, 98);
    const registros = datos.slice(1).map(fila => fila.slice(0, 98));

    const lowerSearch1 = (search1 || "").toLowerCase();
    const lowerSearch2 = (search2 || "").toLowerCase();

    const filtrados = registros.filter(fila => {
      let pasaFiltro1 = true;
      let pasaFiltro2 = true;

      if (lowerSearch1) {
        if (columnaFiltro1 && columnaFiltro1 !== "todos") {
          const colIndex = headers.indexOf(columnaFiltro1);
          if (colIndex !== -1) {
            pasaFiltro1 = fila[colIndex].toLowerCase().includes(lowerSearch1);
          }
        } else {
          pasaFiltro1 = fila.some(celda => celda.toLowerCase().includes(lowerSearch1));
        }
      }

      if (lowerSearch2) {
        if (columnaFiltro2 && columnaFiltro2 !== "todos") {
          const colIndex = headers.indexOf(columnaFiltro2);
          if (colIndex !== -1) {
            pasaFiltro2 = fila[colIndex].toLowerCase().includes(lowerSearch2);
          }
        } else {
          pasaFiltro2 = fila.some(celda => celda.toLowerCase().includes(lowerSearch2));
        }
      }

      return pasaFiltro1 && pasaFiltro2;
    });

    const start = Math.max(filtrados.length - offset - limit, 0);
    const end = filtrados.length - offset;
    const paginados = filtrados.slice(start, end).reverse();

    return {
      headers,
      data: paginados,
      total: filtrados.length
    };
  } catch (error) {
    Logger.log("⚠️ Error en getDatosRegistroCheck: " + error.message);
    return {
      headers: [],
      data: [],
      total: 0,
      error: error.message
    };
  }
}

function getHeadersCheck() {
  const hoja = getCheckSpreadsheet().getSheetByName("B DATOS");
  // ✅ Hasta columna 98 (CT)
  return hoja.getRange(1, 1, 1, 98).getValues()[0]; 
}

function globalVariables() {
  var spreadsheet = getCheckSpreadsheet();

  return {
    spreadsheetId : spreadsheet.getId(),
    dataRage      : 'B DATOS!A2:O',
    idRange       : 'B DATOS!A2:A',
    lastCol       : 'O',
    sheetID       : '675860866'
  };
}

function updateCell(value) {
  var sheetName = "ACTUAL";
  var cellAddress = "J2";
  var sheet = getCheckSpreadsheet().getSheetByName(sheetName);
  if (sheet) {
    sheet.getRange(cellAddress).setValue(value);
  }
}

function getAccessPasswords() {
  const sheet = getCheckSpreadsheet().getSheetByName('Acceso');
  const colB = sheet.getRange('B2:B').getValues().flat().filter(String);
  const colC = sheet.getRange('C2:C').getValues().flat().filter(String);
  return {
    loginPasswords: colB,
    deletePasswords: colC
  };
}

function readData(spreadsheetId, range) {
  var result = Sheets.Spreadsheets.Values.get(spreadsheetId, range);
  return result.values;
}

function deleteData(ID) { 
  var startIndex = getRowIndexByID(ID);
  
  // ✅ ANTES DE BORRAR: Eliminar fotos asociadas del Drive
  try {
    var sheet = getCheckSpreadsheet().getSheetByName('B DATOS');
    var rowNum = startIndex + 1; // startIndex es 0-based (fila header=0), rowNum es 1-based para getRange
    var rowData = sheet.getRange(rowNum, 1, 1, 98).getValues()[0];
    
    // Recopilar todas las URLs de imágenes de esta fila
    var urlsToDelete = [];
    
    // Columna 11 (K) = Imagen principal
    if (rowData[10]) urlsToDelete.push(String(rowData[10]).trim());
    // Columna 12 (L) = Imagen corrección
    if (rowData[11]) urlsToDelete.push(String(rowData[11]).trim());
    
    // Columnas 69-98 (BQ-CT) = Fotos de observación y subsanación por sección
    for (var col = 68; col < 98; col++) {
      if (rowData[col]) urlsToDelete.push(String(rowData[col]).trim());
    }
    
    // Eliminar cada archivo del Drive
    urlsToDelete.forEach(function(url) {
      if (!url || url === '' || url === 'NA') return;
      try {
        var fileId = '';
        // Formato: https://lh5.googleusercontent.com/d/FILE_ID
        if (url.includes('googleusercontent.com/d/')) {
          fileId = url.split('/d/')[1].split(/[?#\/]/)[0];
        }
        // Formato: https://drive.google.com/file/d/FILE_ID/view
        else if (url.includes('drive.google.com/file/d/')) {
          fileId = url.split('/d/')[1].split('/')[0];
        }
        // Formato: https://drive.google.com/open?id=FILE_ID
        else if (url.includes('id=')) {
          fileId = url.split('id=')[1].split('&')[0];
        }
        
        if (fileId) {
          DriveApp.getFileById(fileId).setTrashed(true);
          Logger.log("🗑️ Foto eliminada: " + fileId);
        }
      } catch (e) {
        Logger.log("⚠️ No se pudo eliminar foto: " + url + " - " + e.message);
      }
    });
    
    Logger.log("✅ " + urlsToDelete.filter(u => u && u !== '' && u !== 'NA').length + " fotos procesadas para eliminación del check " + ID);
  } catch (e) {
    Logger.log("⚠️ Error al eliminar fotos del Drive: " + e.message);
  }
  
  // ✅ BORRAR LA FILA
  var deleteRange = {
    "sheetId"     : globalVariables().sheetID,
    "dimension"   : "ROWS",
    "startIndex"  : startIndex,
    "endIndex"    : startIndex + 1
  };
  
  var deleteRequest = [{"deleteDimension":{"range":deleteRange}}];
  Sheets.Spreadsheets.batchUpdate({"requests": deleteRequest}, globalVariables().spreadsheetId);
}

function getRowIndexByID(id) {
  if(id) {
    var idList = readData(globalVariables().spreadsheetId, globalVariables().idRange);
    for(var i = 0; i < idList.length; i++) {
      if(id == idList[i][0]) {
        var rowIndex = parseInt(i + 1);
        return rowIndex;
      }
    }
  }
}

function setStatusCheck(){
  let sst = getCheckSpreadsheet().getSheetByName('B DATOS')
  let totalCheck1 = sst.getRange("T1").getValue();
  let totalCheck2 = sst.getRange("U1").getValue();

  return[totalCheck1, totalCheck2]
}

function getURL() {
  return ScriptApp.getService().getUrl();
}

function searchData(obj) {
  const sheet = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const dataToSearch = sheet.getRange(1, 1, lastRow, 3).getDisplayValues();

  const output = [];

  for (let i = 0; i < dataToSearch.length; i++) {
    if (dataToSearch[i].includes(obj.ad3)) {
      output.push(allData[i]);
    }
  }

  return output;
}

function saveDataCheck(obj) {
  try {
    var sheet = getCheckSpreadsheet().getSheetByName('B DATOS');
    var folder = DriveApp.getFolderById(folderimgcheck);
    var imageUrl = '';

    var idEquipo = String(obj.ad4).trim();

    // 1) Subir imagen si existe
    if (obj.imageData) {
      var imageData = Utilities.base64Decode(obj.imageData.split(',')[1]);
      var blob = Utilities.newBlob(imageData, MimeType.PNG, obj.ad3 + ".png");
      var file = folder.createFile(blob);
      var fileId = file.getId();
      imageUrl = "https://lh5.googleusercontent.com/d/" + fileId;
    } else {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var idColValues = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
        var imgColValues = sheet.getRange(2, 10, lastRow - 1, 1).getValues();

        for (var i = idColValues.length - 1; i >= 0; i--) {
          var idCelda = String(idColValues[i][0]).trim();
          var urlCelda = imgColValues[i][0];
          if (idCelda === idEquipo && urlCelda) {
            imageUrl = urlCelda;
            break;
          }
        }
      }
    }

    // 2) Registrar fila en "B DATOS"
    var timestamp = Math.floor(new Date().getTime() / 1000);
    var newValue = timestamp;
    var status = obj.checked.includes("No") ? "Abierto" : "Conforme";

    // Cálculo de cumplimiento
    var inventarioSheet = getCheckSpreadsheet().getSheetByName('INVENTARIO');
    var inventarioData = inventarioSheet.getRange("A2:N").getValues();

    var diasFrecuencia = null;

    for (var i = 0; i < inventarioData.length; i++) {
      var idInventario = String(inventarioData[i][4]).trim();
      if (idInventario === idEquipo) {
        diasFrecuencia = parseInt(inventarioData[i][13], 10);
        break;
      }
    }

    var cumplimiento = "ID no encontrado";

    if (diasFrecuencia !== null && diasFrecuencia !== "") {
      if (Number(diasFrecuencia) === 0) {
        cumplimiento = "Única vez";
      } else {
        var datos = sheet.getRange("C2:I" + sheet.getLastRow()).getValues();
        var ultimaFecha = null;

        for (var i = datos.length - 1; i >= 0; i--) {
          var idDato = String(datos[i][0]).trim();
          var fechaDato = datos[i][6];

          if (idDato === idEquipo && fechaDato instanceof Date) {
            ultimaFecha = new Date(fechaDato);
            break;
          }
        }

        if (!ultimaFecha) {
          cumplimiento = "Primera vez";
        } else {
          var fechaLimite = new Date(ultimaFecha);
          fechaLimite.setDate(fechaLimite.getDate() + Number(diasFrecuencia));
          var hoy = new Date();
          cumplimiento = (hoy > fechaLimite) ? "No cumple" : "Cumple";
        }
      }
    }

    var checked = "";
    if (obj.checked && typeof obj.checked === 'string') {
      checked = obj.checked;
    } else if (Array.isArray(obj.checked)) {
      checked = obj.checked.join(",");
    }

    // ✅ GUARDAR 15 COLUMNAS (A-O)
    var rowData = [
      newValue,
      obj.ad1,
      obj.ad3,
      obj.ad4,
      obj.ad2,
      obj.ad9,
      obj.ad6,
      obj.ad5,
      obj.ad7,
      new Date(),
      imageUrl,
      "",
      status,
      cumplimiento,
      checked
    ];

    sheet.appendRow(rowData);

    // ✅ GUARDAR OBSERVACIONES POR SECCIÓN EN COLUMNAS BQ-CT (69-98)
    var lastRow = sheet.getLastRow();
    agregarFotosSeccion(obj, lastRow, folder);

    setFormula();

    if (obj.checked.includes("No")) {
      sendChecklistEmail(obj, newValue, imageUrl, status);
    }

    return { 
      columnAValue: newValue,
      success: true,
      clearForm: true
    };

  } catch (error) {
    Logger.log("Error en saveDataCheck: " + error.message);
    return {
      success: false,
      error: error.message,
      clearForm: false
    };
  }
}

// ✅ FUNCIÓN MODIFICADA: Columnas BQ-CT (69-98), 3 columnas por sección
// Col 69,72,75... = Foto/Collage
// Col 70,73,76... = Foto Subsanación (vacío, se llenará después)
// Col 71,74,77... = Comentario observación
function agregarFotosSeccion(obj, fila, folder) {
  var sheet = getCheckSpreadsheet().getSheetByName('B DATOS');
  
  // Parsear sectionObs del HTML (JSON string)
  var sectionObsList = [];
  if (obj.sectionObs) {
    try {
      sectionObsList = JSON.parse(obj.sectionObs);
    } catch (e) {
      Logger.log("⚠️ Error parseando sectionObs: " + e.message);
      return;
    }
  }
  
  if (sectionObsList.length === 0) return;
  
  sectionObsList.forEach(function(obs) {
    // obs.section viene como 0-based desde el HTML, convertir a 1-based
    var seccion = obs.section + 1;
    if (seccion < 1 || seccion > 10) return;
    
    var colFoto = 69 + (seccion - 1) * 3;         // BQ=69, BT=72, BW=75...
    var colSubsanacion = 70 + (seccion - 1) * 3;  // BR=70, BU=73, BX=76... (vacío)
    var colComentario = 71 + (seccion - 1) * 3;    // BS=71, BV=74, BY=77...
    
    // 📸 Guardar foto/collage si existe
    if (obs.image && obs.image !== '') {
      try {
        var base64Data = obs.image.split(',')[1];
        var imageData = Utilities.base64Decode(base64Data);
        var blob = Utilities.newBlob(imageData, MimeType.JPEG, obj.ad3 + "_obs_sec" + seccion + ".jpg");
        var file = folder.createFile(blob);
        var fileId = file.getId();
        var photoUrl = "https://lh5.googleusercontent.com/d/" + fileId;
        
        sheet.getRange(fila, colFoto).setValue(photoUrl);
        Logger.log("✅ Foto observación sección " + seccion + " guardada en columna " + colFoto);
      } catch (error) {
        Logger.log("❌ Error foto sección " + seccion + ": " + error.message);
      }
    }
    
    // 📝 Guardar comentario si existe
    if (obs.comment && obs.comment.trim() !== '') {
      sheet.getRange(fila, colComentario).setValue(obs.comment.trim());
      Logger.log("✅ Comentario sección " + seccion + " guardado en columna " + colComentario);
    }
    
    // ✅ Columna de subsanación queda vacía (se llenará en seguimiento)
  });
}

// ✅ FUNCIÓN: Guardar seguimiento (fotos de levantamiento/subsanación)
function saveDataCheckSeguimiento(objData) {
  try {
    const checkId = objData.checkId;
    const sheet = getCheckSpreadsheet().getSheetByName('B DATOS');
    const folder = DriveApp.getFolderById(folderimgcheck);
    
    if (!checkId) {
      throw new Error("ID del check no proporcionado");
    }
    
    // Buscar la fila del check
    const data = sheet.getRange("A2:A").getValues();
    let rowIndex = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == checkId) {
        rowIndex = i + 2;
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error("Check no encontrado");
    }
    
    // ✅ GUARDAR FOTOS DE SUBSANACIÓN en columnas 70,73,76... (BR,BU,BX...)
    for (let seccion = 1; seccion <= 10; seccion++) {
      const fotoKey = 'fotoLevantamiento-' + seccion;
      
      if (objData[fotoKey] && objData[fotoKey] !== '') {
        try {
          const imageData = Utilities.base64Decode(objData[fotoKey].split(',')[1]);
          const blob = Utilities.newBlob(imageData, MimeType.PNG, 'levantamiento_' + checkId + '_seccion' + seccion + '.png');
          const file = folder.createFile(blob);
          const fileId = file.getId();
          const photoUrl = "https://lh5.googleusercontent.com/d/" + fileId;
          
          // Columna subsanación: 70 + (seccion-1)*3
          const colSubsanacion = 70 + (seccion - 1) * 3;
          sheet.getRange(rowIndex, colSubsanacion).setValue(photoUrl);
          
          Logger.log("✅ Foto subsanación sección " + seccion + " guardada en columna " + colSubsanacion);
        } catch (error) {
          Logger.log("❌ Error guardando foto subsanación sección " + seccion + ": " + error.message);
        }
      }
    }
    
    // ✅ CALCULAR Y ACTUALIZAR ESTADO AUTOMÁTICO
    const estadoActualizado = calcularEstadoCheck(checkId, sheet, rowIndex);
    sheet.getRange(rowIndex, 13).setValue(estadoActualizado);
    
    return {
      success: true,
      checkId: checkId,
      estado: estadoActualizado,
      mensaje: "Seguimiento guardado correctamente"
    };
    
  } catch (error) {
    Logger.log("❌ Error en saveDataCheckSeguimiento: " + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

// ✅ Calcular estado automático del check
function calcularEstadoCheck(checkId, sheet, rowIndex) {
  let totalObservaciones = 0;
  let observacionesLevantadas = 0;
  
  for (let seccion = 1; seccion <= 10; seccion++) {
    const colFoto = 69 + (seccion - 1) * 3;           // Foto observación
    const colSubsanacion = 70 + (seccion - 1) * 3;    // Foto subsanación
    
    const fotoOriginal = sheet.getRange(rowIndex, colFoto).getValue();
    const fotoSubsanacion = sheet.getRange(rowIndex, colSubsanacion).getValue();
    
    if (fotoOriginal && fotoOriginal.toString().trim() !== '') {
      totalObservaciones++;
      if (fotoSubsanacion && fotoSubsanacion.toString().trim() !== '') {
        observacionesLevantadas++;
      }
    }
  }
  
  let estado;
  if (totalObservaciones === 0) {
    estado = "Conforme";
  } else if (observacionesLevantadas === 0) {
    estado = "Abierto";
  } else if (observacionesLevantadas < totalObservaciones) {
    estado = "En Proceso";
  } else {
    estado = "Cerrado";
  }
  
  Logger.log("✅ Estado calculado: " + estado + " (" + observacionesLevantadas + "/" + totalObservaciones + ")");
  return estado;
}

function sendChecklistEmail(obj, newValue, imageUrl, status) {
  const ss = getCheckSpreadsheet();
  
  const menuSheet = ss.getSheetByName('MENÚ');
  const recipient = menuSheet.getRange("B24").getValue().trim();
  if (!recipient) return;

  const itemsSheet = ss.getSheetByName('CHECK LIST');
  const lastRow = itemsSheet.getLastRow();
  const itemsData = itemsSheet.getRange(1, 1, lastRow, 3).getValues();

  const equipo = obj.ad3;

  const items = itemsData
    .filter(row => row[1] === equipo)
    .map(row => row[2]);

  const compList = obj.checked.split(',').map(c => c.trim());

  const itemsHtml = compList.map((valor, i) => {
    const item = items[i];
    if (!item) return '';
    if (valor === 'No') return `<div style="color:#d9534f;"><b>X</b> ${i+1}. ${item}</div>`;
    if (valor === 'Si') return `<div style="color:#5cb85c;"><b>✓</b> ${i+1}. ${item}</div>`;
    return `<div style="color:#0275d8;"><b>O</b> ${i+1}. ${item}</div>`;
  }).join('');

  const subject = "⚠️ Alerta: " + equipo + " No conforme";

  const body = `
    <div style="font-family: Arial; max-width:700px; margin:auto; padding:20px; background:#f9f9f9; border:1px solid #ccc; border-radius:8px;">
      <h2 style="color:#d9534f;">🛑 Alerta Check List – No Conforme</h2>
      <p>Estimado equipo,</p>
      <p>Se ha registrado una lista de verificación con observaciones:</p>
      <table style="width:100%; font-size:14px; margin-top:15px;">
        <tr><td><b>ID</b></td><td>${newValue}</td></tr>
        <tr><td><b>Empresa</b></td><td>${obj.ad1}</td></tr>
        <tr><td><b>Equipo</b></td><td>${equipo}</td></tr>
        <tr><td><b>Código/Placa</b></td><td>${obj.ad4}</td></tr>
        <tr><td><b>Área</b></td><td>${obj.ad2}</td></tr>
        <tr><td><b>Inspector</b></td><td>${obj.ad6}</td></tr>
        <tr><td><b>Proceso</b></td><td>${obj.ad9}</td></tr>
        <tr><td><b>Lugar</b></td><td>${obj.ad5}</td></tr>
        <tr><td><b>Plan de Acción</b></td><td>${obj.ad7}</td></tr>
        <tr><td><b>Fecha</b></td><td>${new Date().toLocaleString()}</td></tr>
        <tr><td><b>Estado</b></td><td style="color:${status === 'Abierto' ? '#d9534f' : '#5cb85c'};"><b>${status}</b></td></tr>
      </table>
      ${ imageUrl ? `<div style="margin-top:20px;"><b>Imagen registrada:</b><br><img src="${imageUrl}" style="max-width:100%; border-radius:4px;"></div>` : '' }
      <div style="margin-top:25px;">
        <h3 style="color:#0275d8;">Lista de Verificación</h3>
        <div style="font-size:14px;">${itemsHtml}</div>
      </div>
      <p style="margin-top:25px;">Revisar y tomar las acciones correspondientes.</p>
      <hr style="margin-top:30px;">
      <p style="font-size:12px; color:#666;">Mensaje generado automáticamente.</p>
    </div>
  `;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: body
  });
}

function setFormula() {
  var sheet = getCheckSpreadsheet().getSheetByName('B DATOS');
  var lastRow = sheet.getLastRow();

  var rangeToCopy = sheet.getRange(lastRow-1, 16);
  rangeToCopy.copyTo(sheet.getRange(lastRow, 16));

  var rangeToCopy = sheet.getRange(lastRow-1, 17);
  rangeToCopy.copyTo(sheet.getRange(lastRow, 17));

  var rangeToCopy = sheet.getRange(lastRow-1, 18);
  rangeToCopy.copyTo(sheet.getRange(lastRow, 18));

  var rangeToCopy = sheet.getRange(lastRow-1, 19);
  rangeToCopy.copyTo(sheet.getRange(lastRow, 19));
}

function getDataCheckList(user) {
  const sheet = getCheckSpreadsheet().getSheetByName("HISTORIAL");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 6).getDisplayValues();

  const result = data.filter(r => r[4] === user.ad3 && r[5] === user.ad4);
  return result;
}

function getDropDownarray() {
  const sheet = getCheckSpreadsheet().getSheetByName("INVENTARIO");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 18).getDisplayValues();

  const filteredData = data.slice(2).map(row => 
    row.slice(1).map(cell => String(cell).trim().replace(/\s+/g, ' '))
  );

  const result = filteredData.filter(row => {
    const estado = row[14].toLowerCase();
    return estado !== "retirado";
  });

  return result;
}

function getAdditionalInfoByValue(value) {
  const sheet = getCheckSpreadsheet().getSheetByName("INVENTARIO");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 14).getDisplayValues();

  for (let i = 2; i < data.length; i++) {
    if (data[i][4] == value) {
      return {
        message1: data[i][7],
        message2: data[i][13]
      };
    }
  }

  return null;
}

function muatData() { 
  var datasheet = getCheckSpreadsheet().getSheetByName("ACTUAL"); 
  var mydata = datasheet.getRange(4,1, datasheet.getLastRow()-3,5).getValues();  
   mydata = mydata.filter(row => row.some(cell => cell !== ''));
  var kolomdata = 0;  
  
 for(var i = 0; i < mydata.length; i++){          
        
        var data = new Date(mydata[i][kolomdata]) ;      

          data.setDate(data.getDate());

        var d = data.getDate();
        var m = data.getMonth() + 1;
        var a = data.getFullYear();
        
        if(d < 10){
           var d = "0" + d;
        }

        if(m < 10){
           var m = "0" + m;
        }
}

 return mydata;
}

function getColumnAValue() {
    var sheet = getCheckSpreadsheet().getSheetByName('B DATOS');
    var lastRow = sheet.getLastRow();
    while (lastRow > 0 && sheet.getRange(lastRow, 1).getValue() === '') {
        lastRow--;
    }

    var columnAValue = sheet.getRange(lastRow, 1).getValue() + 1;
    return { 
      columnAValue: columnAValue,
      success: true,
      clearForm: true
    };
}

function getPdfUrl(columnAValue) {
    var pdfSheet = getCheckSpreadsheet().getSheetByName('FORMATO');
    pdfSheet.getRange('D5').setValue(columnAValue);

    var links = setIDAndGetLinks(columnAValue);

    return links.pdfUrl;
}

function setIDAndGetLinks(recordId) {
    var sheet = getCheckSpreadsheet().getSheetByName('FORMATO');
    sheet.getRange('D5').setValue(recordId);
    
    SpreadsheetApp.flush();
    Utilities.sleep(50);

    var sheetId = sheet.getSheetId();
    var url = getCheckSpreadsheet().getUrl().replace(/edit$/, '');
    
    var exportPdfUrl = url + 'export?format=pdf&gid=' + sheetId + '&range=A1:D55';
    var token = ScriptApp.getOAuthToken();
    var responsePdf = UrlFetchApp.fetch(exportPdfUrl, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });
    var blobPdf = responsePdf.getBlob().setName('PDF_RAC_' + recordId + '.pdf');
    var folder = DriveApp.getFolderById(folderpdfcheck);
    var filePdf = folder.createFile(blobPdf);
    filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var pdfUrl = filePdf.getUrl();
    return {
        pdfUrl: pdfUrl
    };
}

function generarPDF(recordId) {
  var sheet = getCheckSpreadsheet().getSheetByName('FORMATO');
  sheet.getRange('D5').setValue(recordId);
  
  SpreadsheetApp.flush();
  Utilities.sleep(50);

  var sheetId = sheet.getSheetId();
  var url = getCheckSpreadsheet().getUrl().replace(/edit$/, '');
  
  var exportpdflink = url + 'export?format=pdf&gid=' + sheetId + '&range=A1:D55';
  var token = ScriptApp.getOAuthToken();
  var responsePdf = UrlFetchApp.fetch(exportpdflink, {
      headers: {
          'Authorization': 'Bearer ' + token
      }
  });
  var blobPdf = responsePdf.getBlob().setName('PDF_RAC_' + recordId + '.pdf');
  var folder = DriveApp.getFolderById(folderpdfcheck);
  var filePdf = folder.createFile(blobPdf);
  filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var pdflink = filePdf.getUrl();
  return {
      pdflink: pdflink
  };
}

function getData2() {
  const sheet = getCheckSpreadsheet().getSheetByName("B DATOS");
  const lastRow = sheet.getLastRow();
  // ✅ Hasta columna 98 (CT)
  const data = sheet.getRange(1, 1, lastRow, 98).getDisplayValues(); 
  
  console.log(data);
  return data;
}

function getItemsData() {
  const sheet = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 3).getDisplayValues();
  return data;
}

function uploadImageToDrive(fileData, fileName) {
  var folder = DriveApp.getFolderById(folderimgcheck);
  var blob = Utilities.newBlob(Utilities.base64Decode(fileData.split(',')[1]), 'image/png', fileName);
  var file = folder.createFile(blob);
  var fileUrl = "https://lh5.googleusercontent.com/d/" + file.getId();
  return fileUrl;
}

function updateData(updatedData) {
  try {
    const sheet = getCheckSpreadsheet().getSheetByName("B DATOS");
    if (!sheet) throw new Error('La hoja "B DATOS" no existe.');

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const idToFind = updatedData[0];

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == idToFind) {
        const targetRow = i + 2;
        const updateValues = [updatedData.slice(1, 15)];
        sheet.getRange(targetRow, 2, 1, updateValues[0].length).setValues(updateValues);
        break;
      }
    }
  } catch (error) {
    Logger.log("Error en updateData: " + error.message);
    return {
      success: false,
      error: error.message,
      clearForm: false
    };
  }
}

function analizarImagenConLista(base64Image, mimeType, listaItems) {
  const prompt = `
A continuación se presenta una imagen que muestra una situación de campo. Evalúa la imagen y determina para cada ítem de la siguiente lista si la condición se cumple (Sí), no se cumple (No), o no aplica (NA). Solo responde con un JSON en el siguiente formato:

[
  {"item": "Nombre del ítem", "respuesta": "Sí"},
  {"item": "Nombre del ítem", "respuesta": "No"},
  ...
]

Lista de ítems:
${listaItems.join('\n')}
`;

  const requestBody = {
    contents: [
      {
        parts: [
          {
            inline_data: {
              mime_type: mimeType,
              data: base64Image
            }
          },
          {
            text: prompt
          }
        ]
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(requestBody)
  };

  try {
    const response = UrlFetchApp.fetch(geminiUrl, options);
    const json = JSON.parse(response.getContentText());
    const text = json?.candidates?.[0]?.content?.parts?.[0]?.text ?? "";

    const jsonStart = text.indexOf('[');
    const jsonEnd = text.lastIndexOf(']');
    if (jsonStart !== -1 && jsonEnd !== -1) {
      const jsonString = text.substring(jsonStart, jsonEnd + 1);
      return JSON.parse(jsonString);
    }

    return [];
  } catch (error) {
    console.error("Error en analizarImagenConLista:", error);
    return [];
  }
}

function obtenerDatosChecklist() {
  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const lastRow = hoja.getLastRow();
  if (lastRow < 1) return { headersCheck: [], filas: [] };

  const datos = hoja.getRange(1, 1, lastRow, 3).getValues();
  const datosComoTexto = datos.map(fila => fila.map(celda => String(celda || "")));

  return {
    headersCheck: datosComoTexto[0],
    filas: datosComoTexto.slice(1)
  };
}

function obtenerOpcionesInventario() {
  const hoja = getSpreadsheetPersonal().getSheetByName("LISTAS");
  const valores = hoja.getRange("M2:M" + hoja.getLastRow()).getValues().flat();
  return [...new Set(valores.filter(v => v))];
}

function agregarChecklist(data) {
  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  if (!data) return;

  const varias = Array.isArray(data[0]) && data.length > 0;
  const filas = varias ? data : [data];

  const lastRow = hoja.getLastRow();
  let maxId = 0;
  if (lastRow >= 2) {
    const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues().flat()
      .map(v => Number(v))
      .filter(n => !isNaN(n) && isFinite(n));
    if (ids.length) maxId = Math.max.apply(null, ids);
  }

  let nextId = maxId + 1;
  const filasConId = filas.map(row => {
    const r = row.slice();
    if (r.length === 0) {
      return null;
    }
    if (r[0] === undefined || r[0] === null || r[0] === "" || (typeof r[0] === "number" && isNaN(r[0]))) {
      r[0] = nextId++;
    } else {
      const maybeNum = Number(r[0]);
      if (!isNaN(maybeNum)) {
        r[0] = maybeNum;
        if (maybeNum >= nextId) nextId = maybeNum + 1;
      }
    }
    if (r.length < 3) {
      while (r.length < 3) r.push("");
    }
    return r;
  }).filter(Boolean);

  if (filasConId.length === 0) return;

  const startRow = hoja.getLastRow() + 1;
  hoja.getRange(startRow, 1, filasConId.length, filasConId[0].length).setValues(filasConId);
  return { success: true, clearForm: true };
}

function eliminarChecklist(rowIndex) {
  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  hoja.deleteRow(rowIndex + 2);
  return { success: true, clearForm: true };
}

function obtenerItemsPorEquipo(equipo) {
  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  const datos = hoja.getRange(2, 1, lastRow - 1, 3).getValues();

  return datos
    .filter(r => String(r[1]).trim() === String(equipo).trim())
    .map(r => ({ id: Number(r[0]) || 0, item: String(r[2] || "").trim() }));
}

function actualizarChecklistPorEquipo(equipo, itemsJson) {
  if (!equipo) throw new Error("Equipo no especificado.");

  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const nuevosItems = JSON.parse(itemsJson || "[]");
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return;

  const datos = hoja.getRange(2, 1, lastRow - 1, 3).getValues()
    .map((r, i) => ({
      id: Number(r[0]) || 0,
      eq: String(r[1]).trim(),
      item: String(r[2]).trim(),
      fila: i + 2
    }))
    .filter(r => r.eq === String(equipo).trim());

  const existentesPorID = new Map(datos.map(r => [r.id, r]));
  const usadosIDs = new Set();
  const eliminaciones = [];

  nuevosItems.forEach(n => {
    if (n.id && existentesPorID.has(n.id)) {
      const filaExistente = existentesPorID.get(n.id);
      usadosIDs.add(n.id);
      if (n.item !== filaExistente.item) {
        hoja.getRange(filaExistente.fila, 3).setValue(n.item);
      }
    }
  });

  datos.forEach(r => {
    if (!usadosIDs.has(r.id)) eliminaciones.push(r.fila);
  });
  eliminaciones.sort((a, b) => b - a).forEach(f => hoja.deleteRow(f));

  const nuevos = nuevosItems.filter(n => !n.id || !existentesPorID.has(n.id));
  if (nuevos.length > 0) {
    const nextId = (hoja.getRange(hoja.getLastRow(), 1).getValue() || 0) + 1;
    const registros = nuevos.map((n, i) => [nextId + i, equipo, n.item]);
    hoja.getRange(hoja.getLastRow() + 1, 1, registros.length, 3).setValues(registros);
  }

  return {
    status: "ok",
    modificados: usadosIDs.size,
    agregados: nuevos.length,
    eliminados: eliminaciones.length,
    success: true,
    clearForm: true
  };
}

function obtenerInventarioServerSide(offset = 0, limit = 30, terminoBusqueda = "") {
  const hoja = getCheckSpreadsheet().getSheetByName('INVENTARIO');
  const headers = hoja.getRange(2, 1, 1, hoja.getLastColumn()).getValues()[0];
  const ultimaFila = hoja.getLastRow();

  const totalFilas = ultimaFila - 2;
  const filasLeidas = hoja.getRange(3, 1, totalFilas, hoja.getLastColumn()).getValues();
  const fechaCols = [13, 16, 17];

  let datos = filasLeidas.reverse();

  if (terminoBusqueda && terminoBusqueda.trim() !== "") {
    const filtro = terminoBusqueda.toLowerCase();
    datos = datos.filter(fila =>
      fila.some(celda => String(celda).toLowerCase().includes(filtro))
    );
  }

  const paginados = datos.slice(offset, offset + limit);

  const formateados = paginados.map(fila =>
    fila.map((celda, i) => {
      if (fechaCols.includes(i) && celda instanceof Date) {
        return Utilities.formatDate(celda, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      return celda;
    })
  );

  return {
    headersInvet: headers,
    filas: formateados,
    total: datos.length
  };
}

function agregarInventario(data) {
  const hoja = getCheckSpreadsheet().getSheetByName('INVENTARIO');
  data = parsearFechas(data);
  data[0] = hoja.getLastRow() - 1;
  hoja.appendRow(data);
  return { success: true, clearForm: true };
}

function actualizarInventario(data) {
  const hoja = getCheckSpreadsheet().getSheetByName('INVENTARIO');
  data = parsearFechas(data);
  const fila = parseInt(data[0], 10) + 2;
  hoja.getRange(fila, 1, 1, data.length).setValues([data]);
  return { success: true, clearForm: true };
}

function eliminarInventarioPorNum(num) {
  const hoja = getCheckSpreadsheet().getSheetByName('INVENTARIO');
  const fila = parseInt(num, 10) + 2;
  hoja.deleteRow(fila);
  
  const datos = hoja.getRange(3, 1, hoja.getLastRow() - 2, 1).getValues();
  datos.forEach((_, i) => {
    hoja.getRange(i + 3, 1).setValue(i + 1);
  });
  return { success: true, clearForm: true };
}

function parsearFechas(data) {
  const fechaIndices = [13, 16, 17];
  fechaIndices.forEach(i => {
    if (data[i]) {
      const partes = data[i].split("/");
      if (partes.length === 3) {
        data[i] = new Date(`${partes[2]}-${partes[1]}-${partes[0]}`);
      }
    }
  });
  return data;
}

function obtenerEquiposSinChecklist() {
  const hojaInventario = getCheckSpreadsheet().getSheetByName("INVENTARIO");
  const hojaChecklist = getCheckSpreadsheet().getSheetByName("CHECK LIST");

  const limpiarTexto = (texto) => String(texto).trim().toLowerCase();

  const valoresInventario = hojaInventario.getRange("D3:D" + hojaInventario.getLastRow()).getValues().flat();
  const inventarioLimpio = valoresInventario
    .map(limpiarTexto)
    .filter(e => e !== "");

  const valoresChecklist = hojaChecklist.getRange("B2:B" + hojaChecklist.getLastRow()).getValues().flat();
  const checklistLimpio = valoresChecklist
    .map(limpiarTexto)
    .filter(e => e !== "");

  const faltantes = valoresInventario.filter((equipo) => {
    const equipoLimpio = limpiarTexto(equipo);
    return equipoLimpio && !checklistLimpio.includes(equipoLimpio);
  });

  const faltantesUnicos = [...new Set(faltantes.map(limpiarTexto))];

  if (faltantesUnicos.length > 0) {
  }

  return faltantesUnicos;
}

function generarItemsConGemini(base64DataUrl, textoBase, numItems) {

  const hayArchivo = !!base64DataUrl;
  const hayTexto = !!textoBase && textoBase.trim() !== "";

  let prompt = `
Eres un experto en seguridad industrial. Genera una lista de ${
    numItems ? numItems + " " : ""
  }ítems de verificación para un checklist técnico de equipos.

Cada ítem debe tener formato de pregunta breve y clara, con foco en cumplimiento:
Ejemplos:
- "¿Manómetro: Con presión adecuada?"
- "¿Manguera: En buen estado sin daños físicos?"
- "¿Etiqueta de inspección vigente?"
- "¿Válvula principal: Sin fugas visibles?"

`;

  if (hayTexto && hayArchivo) {
    prompt += `
Analiza el documento adjunto y el siguiente texto descriptivo:
"""${textoBase}"""
`;
  } else if (hayArchivo) {
    prompt += `
Analiza el documento adjunto y genera los ítems relevantes del checklist.
`;
  } else if (hayTexto) {
    prompt += `
Basado en este texto o descripción del equipo:
"""${textoBase}"""
`;
  }

  prompt += `
Devuelve un JSON puro con este formato exacto:
[
  {"item": "¿Texto del ítem 1?"},
  {"item": "¿Texto del ítem 2?"},
  ...
]
Solo responde con el JSON, sin explicaciones ni comentarios adicionales.
`;

  const parts = [{ text: prompt }];
  if (hayArchivo) {
    const mimeType = base64DataUrl.match(/^data:(.*?);/)[1];
    const base64 = base64DataUrl.split(",")[1];
    parts.push({ inlineData: { mimeType, data: base64 } });
  }

  const payload = { contents: [{ parts }] };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  const texto = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";

  const inicio = texto.indexOf("[");
  const fin = texto.lastIndexOf("]");
  if (inicio === -1 || fin === -1) throw new Error("Respuesta inválida de Gemini.");

  const limpio = texto.substring(inicio, fin + 1);
  const arr = JSON.parse(limpio);

  return JSON.stringify(arr.map(p => ({ item: p.item || p.pregunta || "" })));
}

function obtenerItemsPorEquipoConSeparadores(equipo) {
  const hoja = getCheckSpreadsheet().getSheetByName("CHECK LIST");
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];
  
  const datos = hoja.getRange(2, 1, lastRow - 1, 3).getValues();
  
  return datos
    .filter(r => String(r[1]).trim() === String(equipo).trim())
    .map(r => {
      const item = String(r[2] || "").trim();
      
      const esNumeroConPunto = /^\d+\.\s*.+$/.test(item);
      const tieneInterrogacion = item.includes('¿') || item.includes('?');
      
      const esTitulo = esNumeroConPunto && !tieneInterrogacion;
      
      const esSubtitulo = !esTitulo;
      
      return {
        id: Number(r[0]) || 0,
        item: item,
        esSeparador: esTitulo,
        seleccionable: esSubtitulo,
        tipo: esTitulo ? 'titulo' : 'subtitulo'
      };
    });
}