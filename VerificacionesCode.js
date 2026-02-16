/* ========================================
   CÓDIGO OPTIMIZADO V2 (AISLADO)
   - INVENTARIO: Caché 10 min
   - CHECK LIST: Caché 10 min
   - B DATOS: Sin caché (transaccional)
   ======================================== */

/* DEFINE GLOBAL VARIABLES V2 */
const SPREADSHEET_IDSCHECK_V2 = {
  check: "1NR4VtBUqO6DkM_rSjNqC8m19-QPjrd_IW1aEmmsUD6U" // ID HOJA PRINCIPAL
};
const folderimgcheck_v2 = "1AnA7M-M7NXVuduuL9ispiLS5mXYvkrb8"; 
let folderpdfcheck_v2 = "15Ofbz86NUhZRuFlWjLUlR6kg1YAY7SXo"; 

// ============================================
// SINGLETON PATTERN: Conexión optimizada V2
// ============================================
let cachedCheck_v2 = null;
function getCheckSpreadsheet_v2() {
  if (!cachedCheck_v2) {
    try {
      cachedCheck_v2 = SpreadsheetApp.openById(SPREADSHEET_IDSCHECK_V2.check);
    } catch (e) {
      console.error("Error conectando sheet V2: " + e.message);
    }
  }
  return cachedCheck_v2;
}

// ============================================
// CACHÉ DE REFERENCIAS A HOJAS V2
// ============================================
const sheetCache_v2 = {};
function getCachedSheet_v2(sheetName) {
  if (!sheetCache_v2[sheetName]) {
    sheetCache_v2[sheetName] = getCheckSpreadsheet_v2().getSheetByName(sheetName);
  }
  return sheetCache_v2[sheetName];
}

// ============================================
// CACHÉ DE INVENTARIO V2 (10 minutos)
// ============================================
let inventarioCache_v2 = null;
let inventarioCacheTime_v2 = 0;
const INVENTARIO_CACHE_TTL_V2 = 600000; 

function getInventarioCached_v2() {
  const now = Date.now();
  
  if (inventarioCache_v2 && (now - inventarioCacheTime_v2) < INVENTARIO_CACHE_TTL_V2) {
    return inventarioCache_v2;
  }
  
  const hoja = getCachedSheet_v2('INVENTARIO');
  const lastRow = hoja.getLastRow();
  
  if (lastRow < 2) {
    inventarioCache_v2 = { data: [], lastRow: 2 };
  } else {
    // Obtenemos hasta la columna 18 (R)
    const data = hoja.getRange(2, 1, lastRow - 2, 8).getValues();
    inventarioCache_v2 = { data, lastRow };
  }
  
  inventarioCacheTime_v2 = now;
  Logger.log(`✅ Caché INVENTARIO V2 actualizado: ${inventarioCache_v2.data.length} registros`);
  return inventarioCache_v2;
}

function invalidarInventarioCache_v2() {
  inventarioCache_v2 = null;
  Logger.log('🔄 Caché INVENTARIO V2 invalidado');
}

// ============================================
// CACHÉ DE CHECK LIST V2 (10 minutos)
// ============================================
let checkListCache_v2 = null;
let checkListCacheTime_v2 = 0;
const CHECKLIST_CACHE_TTL_V2 = 600000; 

function getCheckListCached_v2() {
  const now = Date.now();
  
  if (checkListCache_v2 && (now - checkListCacheTime_v2) < CHECKLIST_CACHE_TTL_V2) {
    return checkListCache_v2;
  }
  
  const hoja = getCachedSheet_v2('CHECK LIST');
  const lastRow = hoja.getLastRow();
  
  if (lastRow < 2) {
    checkListCache_v2 = { headers: [], data: [] };
  } else {
    // CAMBIO: Ahora leemos 5 columnas (A hasta E)
    const allData = hoja.getRange(1, 1, lastRow, 5).getDisplayValues();
    checkListCache_v2 = {
      headers: allData[0],
      data: allData.slice(1)
    };
  }
  
  checkListCacheTime_v2 = now;
  // Logger.log(`✅ Caché CHECK LIST V2 actualizado`);
  return checkListCache_v2;
}

function invalidarCheckListCache_v2() {
  checkListCache_v2 = null;
  Logger.log('🔄 Caché CHECK LIST V2 invalidado');
}

// ============================================
// CACHÉ DE PASSWORDS V2 (5 minutos)
// ============================================
let cachedPasswords_v2 = null;
let passwordsCacheTime_v2 = 0;
const PASSWORDS_CACHE_TTL_V2 = 300000; 

function getAccessPasswords_v2() {
  const now = Date.now();
  if (cachedPasswords_v2 && (now - passwordsCacheTime_v2) < PASSWORDS_CACHE_TTL_V2) {
    return cachedPasswords_v2;
  }
  
  const sheet = getCachedSheet_v2('Acceso');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { loginPasswords: [], deletePasswords: [] };

  const data = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
  const login = [], del = [];
  
  for(let i = 0; i < data.length; i++) {
    if(data[i][0]) login.push(String(data[i][0]));
    if(data[i][1]) del.push(String(data[i][1]));
  }

  cachedPasswords_v2 = { loginPasswords: login, deletePasswords: del };
  passwordsCacheTime_v2 = now;
  return cachedPasswords_v2;
}

/* =========================================================
   FUNCIONES PÚBLICAS (Llamadas desde Frontend V2)
   ========================================================= */

// 1. Obtener Datos para la Tabla Principal (Historial)
function getDatosRegistroCheck_v2(offset, limit, search1, search2, columnaFiltro1, columnaFiltro2) {
  try {
    const hoja = getCachedSheet_v2("B DATOS");
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return { headers: [], data: [], total: 0 };

    const dataRange = hoja.getRange(1, 1, lastRow, 16).getDisplayValues();
    const headers = dataRange[0].slice(0, 16);
    const registros = dataRange.slice(1).reverse();

    const lowerSearch1 = (search1 || "").toLowerCase();
    const lowerSearch2 = (search2 || "").toLowerCase();

    const filtrados = registros.filter(fila => {
      let pasaFiltro1 = true;
      let pasaFiltro2 = true;

      if (lowerSearch1) {
        if (columnaFiltro1 && columnaFiltro1 !== "todos") {
          const colIndex = headers.indexOf(columnaFiltro1);
          if (colIndex !== -1) pasaFiltro1 = fila[colIndex].toLowerCase().includes(lowerSearch1);
        } else {
          pasaFiltro1 = fila.some(celda => celda.toLowerCase().includes(lowerSearch1));
        }
      }

      if (lowerSearch2 && pasaFiltro1) {
        if (columnaFiltro2 && columnaFiltro2 !== "todos") {
          const colIndex = headers.indexOf(columnaFiltro2);
          if (colIndex !== -1) pasaFiltro2 = fila[colIndex].toLowerCase().includes(lowerSearch2);
        } else {
          pasaFiltro2 = fila.some(celda => celda.toLowerCase().includes(lowerSearch2));
        }
      }

      return pasaFiltro1 && pasaFiltro2;
    });

    const paginados = filtrados.slice(offset, offset + limit);

    return {
      headers,
      data: paginados,
      total: filtrados.length
    };
  } catch (error) {
    Logger.log("⚠️ Error en getDatosRegistroCheck_v2: " + error.message);
    return { headers: [], data: [], total: 0, error: error.message };
  }
}

// 2. Obtener Encabezados
function getHeadersCheck_v2() {
  const hoja = getCachedSheet_v2("B DATOS");
  return hoja.getRange(1, 1, 1, 14).getValues()[0];
}

// 3. Eliminar Registro
function deleteData_v2(ID) { 
  if (!ID) return false;

  try {
    const sheet = getCachedSheet_v2("B DATOS");
    const result = sheet.getRange("A:A")
                        .createTextFinder(String(ID))
                        .matchEntireCell(true)
                        .findNext();
    
    if (result) {
      sheet.deleteRow(result.getRow());
      Logger.log("ID V2 " + ID + " eliminado correctamente.");
      return true;
    } else {
      return false;
    }
  } catch (e) {
    Logger.log("Error en deleteData_v2: " + e.message);
    return false;
  }
}

// 4. Totales (Abiertos/Cerrados)
function setStatusCheck_v2(){
  let sst = getCachedSheet_v2('B DATOS');
  let data = sst.getRange("U1:V1").getValues()[0];
  return [data[0], data[1]];
}

// 5. Buscar Items para Checklist (Formulario)
function searchData_v2(obj) {
  const cache = getCheckListCached_v2();
  return cache.data.filter(row => row.length >= 4 && String(row[1]) === String(obj.ad3));
}

// 6. Guardar Nuevo Checklist
// 6. Guardar Nuevo Checklist (AJUSTADO: Sin Plazo, Con Actividad)
function saveDataCheck_v2(obj) {
  try {
    const sheetBD = getCachedSheet_v2('B DATOS');
    const idEquipo = String(obj.ad4).trim();
    const ahora = new Date();
    const timestamp = Math.floor(ahora.getTime() / 1000);
    
    // 1. PROCESAR IMAGEN PRINCIPAL
    let imageUrl = '';
    if (obj.imageData && obj.imageData.includes('base64')) {
      const folder = DriveApp.getFolderById(folderimgcheck_v2);
      const imageData = Utilities.base64Decode(obj.imageData.split(',')[1]);
      const blob = Utilities.newBlob(imageData, MimeType.PNG, `${obj.ad3}_${idEquipo}_V2.png`);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrl = "https://lh5.googleusercontent.com/d/" + file.getId();
    } else {
      // Intentar recuperar imagen histórica si no se sube una nueva
      const historicoBD = sheetBD.getRange(2, 1, Math.max(1, sheetBD.getLastRow() - 1), 11).getValues();
      const registroConImagen = historicoBD.reverse().find(r => String(r[3]).trim() === idEquipo && r[10]);
      imageUrl = registroConImagen ? registroConImagen[10] : '';
    }

    // 2. PROCESAR IMAGEN CORRECCIÓN
    let imgCorrUrl = '';
    if (obj.imgCorreccion && obj.imgCorreccion.includes('base64')) {
      const folder = DriveApp.getFolderById(folderimgcheck_v2);
      const imageData = Utilities.base64Decode(obj.imgCorreccion.split(',')[1]);
      const blob = Utilities.newBlob(imageData, MimeType.PNG, `CORR_${timestamp}_V2.png`);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imgCorrUrl = "https://lh5.googleusercontent.com/d/" + file.getId();
    }

    // 3. DEFINIR ESTADO
    const status = String(obj.checked).includes("No") ? "Abierto" : "Conforme";
    
    // 4. PREPARAR FILA (Columna N ahora es ACTIVIDAD)
    const rowData = [
      timestamp,     // A (0)
      obj.ad1,       // B (1)
      obj.ad3,       // C (2)
      idEquipo,      // D (3)
      obj.ad2,       // E (4)
      obj.ad9,       // F (5)
      obj.ad6,       // G (6)
      obj.ad5,       // H (7)
      obj.ad7,       // I (8)
      ahora,         // J (9)
      imageUrl,      // K (10)
      imgCorrUrl,    // L (11)
      status,        // M (12)
      obj.actividad, // N (13) <-- AQUÍ VA LA ACTIVIDAD (Antes Plazo)
      obj.riesgoCheck, // O (14) 
      obj.checked    // P (15)
    ];

    sheetBD.appendRow(rowData);
    setFormula_v2();

    // 5. ENVÍO DE ALERTA
    if (status === "Abierto") {
      sendChecklistEmail_v2(obj, timestamp, imageUrl, status);
    }

    return { columnAValue: timestamp };

  } catch (error) {
    Logger.log("Error en saveDataCheck_v2: " + error.message);
    throw new Error("Error V2: " + error.message);
  }
}

function sendChecklistEmail_v2(obj, newValue, imageUrl, status) {
  const menuSheet = getCachedSheet_v2('MENÚ');
  const recipient = menuSheet.getRange("B24").getValue().trim();
  if (!recipient) return;

  const checkCache = getCheckListCached_v2();
  
  // 1. OBTENER ITEMS (PREGUNTAS)
  // Filtramos por el nombre del checklist (Columna B / Índice 1)
  // Obtenemos la pregunta (Columna D / Índice 3)
  const items = checkCache.data
    .filter(row => String(row[1]) === String(obj.ad3))
    .map(row => row[3]); 

  // 2. PARSEAR RESPUESTAS
  const compList = obj.checked.split(',').map(c => c.trim());

  // 3. GENERAR HTML DEL CUERPO
  const itemsHtml = compList.map((valor, i) => {
    const item = items[i] || "Ítem desconocido";
    
    // Configuración por defecto (Neutro/Azul)
    let style = "color:#0275d8;"; 
    let icon = "<b>●</b>"; 
    let textoValor = valor;

    // LÓGICA DE SEMÁFORO (Soporta 1-5, Si/No y Texto)
    if (valor === '1' || valor === '2' || valor === 'No') { 
        // MALO (Rojo)
        style = "color:#d9534f; font-weight:bold;"; 
        icon = "<b>✕</b>";
    } else if (valor === '4' || valor === '5' || valor === 'Si') { 
        // BUENO (Verde)
        style = "color:#28a745;"; 
        icon = "<b>✓</b>";
    } else if (valor === 'NA' || valor === '3') {
        // REGULAR / NA (Gris/Azul)
        style = "color:#6c757d;";
        icon = "<b>○</b>";
    } else {
        // TEXTO (Dejar en azul informativo)
        icon = "<b>📝</b>";
    }

    return `
      <div style="border-bottom:1px solid #eee; padding:5px 0; font-size:13px;">
        <span style="${style} margin-right:8px; font-size:14px;">${icon}</span>
        <span style="color:#333;">${i+1}. ${item}</span>
        <div style="font-size:11px; color:#666; margin-left:24px; font-style:italic;">Respuesta: ${textoValor}</div>
      </div>`;
  }).join('');

  const subject = "⚠️ Reporte V2: " + obj.ad3 + " - " + status;
  
  const body = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width:600px; margin:auto; background:#ffffff; border:1px solid #e0e0e0; border-radius:8px; overflow:hidden;">
      
      <div style="background:${status === 'Abierto' ? '#d9534f' : '#28a745'}; color:white; padding:15px; text-align:center;">
        <h2 style="margin:0; font-size:20px;">REPORTE ${status.toUpperCase()}</h2>
        <p style="margin:5px 0 0 0; opacity:0.9;">ID: ${newValue} | ${new Date().toLocaleDateString()}</p>
      </div>

      <div style="padding:20px;">
        <table style="width:100%; font-size:13px; color:#444; margin-bottom:20px;">
          <tr><td style="font-weight:bold; width:30%;">Equipo:</td><td>${obj.ad3}</td></tr>
          <tr><td style="font-weight:bold;">ID/Placa:</td><td>${obj.ad4}</td></tr>
          <tr><td style="font-weight:bold;">Inspector:</td><td>${obj.ad6}</td></tr>
          <tr><td style="font-weight:bold;">Ubicación:</td><td>${obj.ad5}</td></tr>
          <tr><td style="font-weight:bold;">Proceso:</td><td>${obj.ad9}</td></tr>
          <tr><td style="font-weight:bold;">Riesgo Crítico:</td><td>${obj.riesgoCheck}</td></tr>
          <tr><td style="font-weight:bold;">Plan Acción:</td><td>${obj.RCRI}</td></tr>
        </table>

        ${imageUrl ? `
        <div style="text-align:center; margin-bottom:20px; border:1px solid #eee; padding:5px; border-radius:5px;">
          <img src="${imageUrl}" style="max-width:100%; height:auto; border-radius:4px;">
          <p style="font-size:11px; color:#999; margin:5px 0 0 0;">Evidencia Principal</p>
        </div>` : ''}

        <hr style="border:0; border-top:1px solid #eee; margin:20px 0;">

        <h3 style="color:#004aad; font-size:16px; margin-bottom:15px;">Detalle de Inspección</h3>
        <div>${itemsHtml}</div>

      </div>
      
      <div style="background:#f8f9fa; padding:10px; text-align:center; font-size:11px; color:#888;">
        Reporte generado automáticamente por Sistema SSOMA V2
      </div>
    </div>
  `;

  MailApp.sendEmail({ to: recipient, subject: subject, htmlBody: body });
}

// 8. Fórmulas
function setFormula_v2() {
  var sheet = getCachedSheet_v2('B DATOS');
  var lastRow = sheet.getLastRow();
  if(lastRow < 2) return;

  var sourceRange = sheet.getRange(lastRow-1, 17, 1, 4); 
  sourceRange.copyTo(sheet.getRange(lastRow, 17));
}

// 9. Historial (Botón Reloj)
function getDataCheckList_v2(user) {
  const sheet = getCachedSheet_v2("HISTORIAL"); // O B DATOS, ajusta según necesidad
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 6).getDisplayValues(); // Ajusta rango si es B DATOS
  return data.filter(r => r[4] === user.ad3 && r[5] === user.ad4);
}

// 10. Listas Dropdown (Cascada)
function getDropDownarray_v2() {
  const cache = getInventarioCached_v2();
  
  const result = [];
  for(let i = 0; i < cache.data.length; i++) {
    if(String(cache.data[i][7]).toLowerCase() !== "retirado") {
      result.push(cache.data[i].slice(1, 12).map(cell => String(cell).trim().replace(/\s+/g, ' ')));
    }
  }
  return result;
}

// 11. Info Adicional
function getAdditionalInfoByValue_v2(value) {
  const cache = getInventarioCached_v2();
  
  for (let i = 0; i < cache.data.length; i++) {
    if (cache.data[i][4] == value) { // Ajusta índice si columna ID cambia
      return {
        message1: cache.data[i][5],
        message2: cache.data[i][6]
      };
    }
  }
  return null;
}

// 12. Obtener Nombre Usuario
function getUserNameFromServer_v2() {
  return Session.getActiveUser().getEmail(); 
}

// 13. Obtener PDF URL
function getPdfUrl_v2(columnAValue) {
  var pdfSheet = getCachedSheet_v2('FORMATO');
  pdfSheet.getRange('E5').setValue(columnAValue);
  var links = setIDAndGetLinks_v2(columnAValue);
  return links.pdfUrl;
}

function setIDAndGetLinks_v2(recordId) {
  var sheet = getCachedSheet_v2('FORMATO');
  sheet.getRange('E5').setValue(recordId);
  
  SpreadsheetApp.flush();

  var sheetId = sheet.getSheetId();
  var url = getCheckSpreadsheet_v2().getUrl().replace(/edit$/, '');
  
  var exportPdfUrl = url + 'export?format=pdf&gid=' + sheetId + '&range=A1:E55';
  var token = ScriptApp.getOAuthToken();
  
  var responsePdf = UrlFetchApp.fetch(exportPdfUrl, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });
  
  if (responsePdf.getResponseCode() !== 200) {
    throw new Error("Error generando PDF V2");
  }

  var blobPdf = responsePdf.getBlob().setName('PDF_RAC_V2_' + recordId + '.pdf');
  var folder = DriveApp.getFolderById(folderpdfcheck_v2);
  var filePdf = folder.createFile(blobPdf);
  filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return { pdfUrl: filePdf.getUrl() };
}

// 14. Items Data
function getItemsData_v2() {
  const cache = getCheckListCached_v2();
  return [cache.headers, ...cache.data];
}

// 15. IA Analizar Imagen
function analizarImagenConLista_v2(base64Image, mimeType, listaItems) {
  if(typeof API_KEY_V2 === 'undefined' || API_KEY_V2 ===  API_KEY) {
    return []; 
  }

  const prompt = `
A continuación se presenta una imagen que muestra la evaluación correcta del llenado del fomulario. Evalúa la imagen y determina para cada ítem de la siguiente lista si la condición se cumple (Sí), no se cumple (No), o no aplica (NA). Solo responde con un JSON en el siguiente formato:
[{"item": "Nombre del ítem", "respuesta": "Sí"}, ...]
Lista de ítems:
${listaItems.join('\n')}
`;

  const requestBody = {
    contents: [{
      parts: [
        { inline_data: { mime_type: mimeType, data: base64Image } },
        { text: prompt }
      ]
    }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;
    const response = UrlFetchApp.fetch(url, options);
    if(response.getResponseCode() !== 200) return [];

    const json = JSON.parse(response.getContentText());
    const text = json?.candidates?.[0]?.content?.parts?.[0]?.text ?? "";

    const jsonStart = text.indexOf('[');
    const jsonEnd = text.lastIndexOf(']');
    if (jsonStart !== -1 && jsonEnd !== -1) {
      return JSON.parse(text.substring(jsonStart, jsonEnd + 1));
    }
    return [];
  } catch (error) {
    console.error("Error en analizarImagenConLista_v2:", error);
    return [];
  }
}

// 16. Actualizar Datos Checklist
// 16. Actualizar Datos Checklist (AJUSTADO: Imágenes + Actividad)
function updateDataChecklistServer_v2(obj) {
  try {
    const sheet = getCachedSheet_v2('B DATOS');
    const idToFind = String(obj.id); 
    
    const finder = sheet.getRange("A:A").createTextFinder(idToFind).matchEntireCell(true);
    const result = finder.findNext();

    if (!result) throw new Error("ID V2 no encontrado: " + idToFind);
    
    const row = result.getRow();
    const folder = DriveApp.getFolderById(folderimgcheck_v2);

    // --- A. IMAGEN PRINCIPAL (Col 11/K) ---
    let finalMainUrl = "";
    if (obj.imageData && obj.imageData.includes("base64")) {
      const imageData = Utilities.base64Decode(obj.imageData.split(',')[1]);
      const blob = Utilities.newBlob(imageData, MimeType.PNG, `${obj.ad3}_${obj.ad4}_MAIN_UPD.png`);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      finalMainUrl = "https://lh5.googleusercontent.com/d/" + file.getId();
    } else {
      finalMainUrl = sheet.getRange(row, 11).getValue();
    }

    // --- B. IMAGEN CORRECCIÓN (Col 12/L) ---
    let finalCorrUrl = "";
    if (obj.imgCorreccion && obj.imgCorreccion.includes("base64")) {
      const imageData = Utilities.base64Decode(obj.imgCorreccion.split(',')[1]);
      const blob = Utilities.newBlob(imageData, MimeType.PNG, `CORR_${obj.ad4}_UPD.png`);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      finalCorrUrl = "https://lh5.googleusercontent.com/d/" + file.getId();
    } else {
      finalCorrUrl = sheet.getRange(row, 12).getValue();
    }

    // --- C. GUARDAR DATOS (Col 14/N es Actividad) ---
    sheet.getRange(row, 6).setValue(obj.ad9);   // Proceso
    sheet.getRange(row, 8).setValue(obj.ad5);   // Lugar
    sheet.getRange(row, 9).setValue(obj.ad7);   // Plan
    sheet.getRange(row, 11).setValue(finalMainUrl); // Imagen Principal
    sheet.getRange(row, 12).setValue(finalCorrUrl); // Imagen Corrección
    sheet.getRange(row, 13).setValue(obj.estado); // Estado
    
    // CAMBIO IMPORTANTE: Guardamos Actividad en lugar de Plazo
    sheet.getRange(row, 14).setValue(obj.actividad); 
    sheet.getRange(row, 15).setValue(obj.riesgoCheck); 
    
    sheet.getRange(row, 16).setValue(obj.checked); // Checks

    return { status: "success" };

  } catch (error) {
    Logger.log("Error V2: " + error.message);
    throw new Error(error.message);
  }
}

//INVENTARIO
// --- CONSTANTES Y CONFIGURACIÓN ---
const NUM_COLS_INVENT_EXTRA = 8; 
const NOMBRE_HOJA_EXTRA = 'INVENTARIO'; // <--- NOMBRE REAL DE TU HOJA

// --- CACHÉ ESPECÍFICO PARA INVENTARIO EXTRA ---
let inventarioExtraCache_v2 = null;
let inventarioExtraCacheTime_v2 = 0;
const INVENTARIO_EXTRA_CACHE_TTL = 600000; 

function getInventarioExtraCached_v2() {
  const now = Date.now();
  
  if (inventarioExtraCache_v2 && (now - inventarioExtraCacheTime_v2) < INVENTARIO_EXTRA_CACHE_TTL) {
    return inventarioExtraCache_v2;
  }
  
  const hoja = getCachedSheet_v2(NOMBRE_HOJA_EXTRA);
  const lastRow = hoja.getLastRow();
  
  // CORRECCIÓN: Si hay menos de 2 filas, significa que solo hay encabezado o está vacía
  if (lastRow < 2) { 
    inventarioExtraCache_v2 = { data: [] };
  } else {
    // CORRECCIÓN: Leemos desde la fila 2 (Datos) hasta el final
    // Fórmula de filas: (Total - FilaInicio + 1) -> (lastRow - 2 + 1) = lastRow - 1
    const data = hoja.getRange(2, 1, lastRow - 1, NUM_COLS_INVENT_EXTRA).getValues();
    inventarioExtraCache_v2 = { data };
  }
  
  inventarioExtraCacheTime_v2 = now;
  Logger.log(`✅ Caché INVENTARIO EXTRA actualizado: ${inventarioExtraCache_v2.data.length} registros`);
  return inventarioExtraCache_v2;
}

function invalidarInventarioExtraCache_v2() {
  inventarioExtraCache_v2 = null;
  Logger.log('🔄 Caché INVENTARIO EXTRA invalidado');
}

// --- FUNCIONES DEL SERVIDOR ---

function obtenerInventarioServerSideExtra(offset = 0, limit = 30, terminoBusqueda = "") {
  const invCache = getInventarioExtraCached_v2(); 
  
  const hoja = getCachedSheet_v2(NOMBRE_HOJA_EXTRA);
  
  // CORRECCIÓN PRINCIPAL: Aquí cambiamos el 2 por el 1 para leer la PRIMERA fila
  const headers = hoja.getRange(1, 1, 1, NUM_COLS_INVENT_EXTRA).getValues()[0];
  
  if (!invCache.data || invCache.data.length === 0) {
    return { headersInvet: headers, filas: [], total: 0 };
  }

  let datos = [...invCache.data].reverse();
  
  let filtrados = datos;
  if (terminoBusqueda.trim()) {
    const filtro = terminoBusqueda.toLowerCase();
    filtrados = datos.filter(fila => 
      fila.some(c => String(c).toLowerCase().includes(filtro))
    );
  }

  const paginados = filtrados.slice(offset, offset + limit);
  
  return { headersInvet: headers, filas: paginados, total: filtrados.length };
}

function agregarInventarioExtra(data) {
  const hoja = getCachedSheet_v2(NOMBRE_HOJA_EXTRA);
  const lastRow = hoja.getLastRow();
  let nuevoId = 1;
  
  // CORRECCIÓN: Ajustado para leer IDs desde la fila 2
  if (lastRow >= 2) { 
    const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const soloNumeros = ids.map(Number).filter(n => !isNaN(n));
    if (soloNumeros.length > 0) {
      nuevoId = Math.max(...soloNumeros) + 1;
    }
  }
  
  data[0] = nuevoId; 
  hoja.appendRow(data);
  
  invalidarInventarioExtraCache_v2(); 
  return { status: "success", id: nuevoId };
}

function actualizarInventarioExtra(data) {
  const hoja = getCachedSheet_v2(NOMBRE_HOJA_EXTRA);
  const idABuscar = String(data[0]); 
  
  const finder = hoja.getRange("A:A").createTextFinder(idABuscar).matchEntireCell(true);
  const result = finder.findNext();
  
  if (result) {
    const filaReal = result.getRow();
    hoja.getRange(filaReal, 1, 1, data.length).setValues([data]);
    
    invalidarInventarioExtraCache_v2(); 
    return { status: "success", fila: filaReal };
  } else {
    throw new Error("No se encontró el registro con ID: " + idABuscar);
  }
}

function eliminarInventarioExtraPorNum(num) {
  const hoja = getCachedSheet_v2(NOMBRE_HOJA_EXTRA);
  
  const finder = hoja.getRange("A:A").createTextFinder(String(num)).matchEntireCell(true);
  const result = finder.findNext();
  
  if(result){
    const fila = result.getRow();
    hoja.deleteRow(fila);
    
    // CORRECCIÓN: Renumerar IDs desde la fila 2
    const lastRow = hoja.getLastRow();
    if(lastRow >= 2) {
      const range = hoja.getRange(2, 1, lastRow - 1, 1);
      const newIds = [];
      for(let i = 0; i < lastRow - 1; i++) newIds.push([i + 1]);
      range.setValues(newIds);
    }
    
    invalidarInventarioExtraCache_v2(); 
  }
}

//CHECK
/* =========================================================
   MÓDULO CHECKLIST EXTRA (5 COLUMNAS) - SIN IA
   Columnas: [0]Num, [1]Herramienta, [2]Categoría, [3]Pregunta, [4]Opciones
   ========================================================= */

const NOMBRE_HOJA_CHECK_EXTRA = 'CHECK LIST';

// --- CACHÉ CHECKLIST EXTRA ---
let checkListExtraCache = null;
const CHECKLIST_EXTRA_TTL = 600000; 
let checkListExtraTime = 0;

function getCheckListExtraCached() {
  const now = Date.now();
  if (checkListExtraCache && (now - checkListExtraTime) < CHECKLIST_EXTRA_TTL) {
    return checkListExtraCache;
  }
  
  const hoja = getCachedSheet_v2(NOMBRE_HOJA_CHECK_EXTRA);
  const lastRow = hoja.getLastRow();
  
  if (lastRow < 2) {
    checkListExtraCache = { headers: [], data: [] };
  } else {
    // Leemos 5 columnas
    const values = hoja.getRange(1, 1, lastRow, 5).getValues();
    checkListExtraCache = {
      headers: values[0], 
      data: values.slice(1) 
    };
  }
  checkListExtraTime = now;
  return checkListExtraCache;
}

function invalidarCheckListExtraCache() {
  checkListExtraCache = null;
}

// --- FUNCIONES CRUD ---

function obtenerDatosChecklistExtra() {
  const cache = getCheckListExtraCached();
  return {
    headersCheck: cache.headers, 
    filas: cache.data
  };
}

function obtenerOpcionesInventarioExtra() {
  try {
    const invCache = getInventarioExtraCached_v2(); 
    const herramientas = invCache.data.map(r => r[3]).filter(h => h);
    return [...new Set(herramientas)];
  } catch(e) { return []; }
}

function eliminarChecklistExtra(rowIndex) {
  const hoja = getCachedSheet_v2(NOMBRE_HOJA_CHECK_EXTRA);
  hoja.deleteRow(rowIndex + 2); 
  invalidarCheckListExtraCache();
}

function obtenerItemsPorEquipoExtra(equipo) {
  const cache = getCheckListExtraCached();
  
  return cache.data
    .filter(r => String(r[1]).trim() === String(equipo).trim())
    .map(r => ({ 
      id: Number(r[0]) || 0, 
      categoria: String(r[2] || "").trim(),
      item: String(r[3] || "").trim(),
      opciones: String(r[4] || "").trim() // Ahora guardará BINARIO, TEXTO o ESCALA
    }));
}

function actualizarChecklistPorEquipoExtra(equipo, itemsJson) {
  if (!equipo) throw new Error("Equipo no especificado.");
  
  const hoja = getCachedSheet_v2(NOMBRE_HOJA_CHECK_EXTRA);
  const nuevosItems = JSON.parse(itemsJson || "[]"); 
  const lastRow = hoja.getLastRow();
  
  let datosDB = [];
  if(lastRow >= 2) {
    datosDB = hoja.getRange(2, 1, lastRow - 1, 5).getValues()
      .map((r, i) => ({
        id: Number(r[0]) || 0,
        eq: String(r[1]).trim(),
        cat: String(r[2]),
        preg: String(r[3]),
        opt: String(r[4]),
        fila: i + 2
      }))
      .filter(r => r.eq === String(equipo).trim());
  }

  const existentesPorID = new Map(datosDB.map(r => [r.id, r]));
  const usadosIDs = new Set();
  const eliminaciones = [];

  nuevosItems.forEach(n => {
    if (n.id && existentesPorID.has(n.id)) {
      const registro = existentesPorID.get(n.id);
      usadosIDs.add(n.id);
      
      if (n.item !== registro.preg || n.categoria !== registro.cat || n.opciones !== registro.opt) {
        hoja.getRange(registro.fila, 3, 1, 3).setValues([[n.categoria, n.item, n.opciones]]);
      }
    }
  });

  datosDB.forEach(r => { if (!usadosIDs.has(r.id)) eliminaciones.push(r.fila); });
  eliminaciones.sort((a, b) => b - a).forEach(f => hoja.deleteRow(f));

  const nuevos = nuevosItems.filter(n => !n.id || !existentesPorID.has(n.id));
  
  if (nuevos.length > 0) {
    const allIds = hoja.getRange("A:A").getValues().flat();
    let maxId = 0;
    allIds.forEach(id => { let n = Number(id); if(!isNaN(n) && n > maxId) maxId = n; });
    
    // Aquí asignamos un valor por defecto solo si viene vacío, pero el frontend enviará uno de los 3.
    const filasAInsertar = nuevos.map((n, i) => [
      maxId + 1 + i, 
      equipo, 
      n.categoria || "General", 
      n.item, 
      n.opciones || "BINARIO" 
    ]);
    
    hoja.getRange(hoja.getLastRow() + 1, 1, filasAInsertar.length, 5).setValues(filasAInsertar);
  }

  invalidarCheckListExtraCache();

  return { status: "ok", modificados: usadosIDs.size, agregados: nuevos.length, eliminados: eliminaciones.length };
}

function agregarChecklistExtra(data) {
  const hoja = getCachedSheet_v2(NOMBRE_HOJA_CHECK_EXTRA); 
  
  const lastRow = hoja.getLastRow();
  let maxId = 0;
  if(lastRow >= 2) {
    const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    ids.forEach(id => { let n = Number(id); if(!isNaN(n) && n > maxId) maxId = n; });
  }

  let nextId = maxId + 1;
  const filasConId = data.map(r => {
    let row = [...r];
    row[0] = nextId++; 
    while(row.length < 5) row.push(""); 
    return row;
  });

  if (filasConId.length > 0) {
    hoja.getRange(lastRow + 1, 1, filasConId.length, 5).setValues(filasConId);
  }
  invalidarCheckListExtraCache();
}

function obtenerEquiposSinChecklistExtra() {
  const invCache = getInventarioExtraCached_v2();
  const checkCache = getCheckListExtraCached();

  const limpiar = (t) => String(t).trim().toLowerCase();
  
  const equiposInventario = [...new Set(invCache.data.map(r => limpiar(r[3])).filter(e => e))];
  const equiposChecklist = new Set(checkCache.data.map(r => limpiar(r[1])).filter(e => e));

  const faltantes = invCache.data
    .filter(r => {
       const nombre = limpiar(r[3]);
       return nombre && !equiposChecklist.has(nombre);
    })
    .map(r => r[3]);

  return [...new Set(faltantes)];
}
