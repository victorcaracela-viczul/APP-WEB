//Credit: https://www.youtube.com/watch?v=CMwzLURK-rQ

function doGet() {
  return HtmlService.createTemplateFromFile('index')
  .evaluate()
  .setTitle('GESTIÓN OP - SSOMA')
  .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const API_KEY = 'AIzaSyDGc1jUaCDs6h4FSBi3ddIAahCRjiZ_JPg';
//const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${API_KEY}`;
const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;
//let personal = SpreadsheetApp.openById("1X2zQSVpj3HkGptI2n5LdLZi4ikT0vGU2mQCnfah2QhQ") 
const folderIdFirma = '1TzV9UlPupxeRyo7l2Vn2nO9mh64WG_Kv'; //GESTION - FIRMA


let cachedPersonal = null;
function getSpreadsheetPersonal() {
  if (!cachedPersonal) {
    cachedPersonal = SpreadsheetApp.openById("1NDDHlTfWxmObgm8JZu5WAnCECB3gU6e_k7o_sFcMrkw"); // HOJA DE CALCULO GESTION PERSONAL
  }
  return cachedPersonal; 
}

function loginData(obj) {
  const hoja = getSpreadsheetPersonal().getSheetByName('PERSONAL');
  const lastRow = hoja.getLastRow();
  
  // Una sola lectura de todas las columnas necesarias
  const data = hoja.getRange(2, 1, lastRow - 1, 18).getValues();
  
  const id = obj.username.toString().toLowerCase() + (obj.password || "");
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const username = (row[1] || "").toString().toLowerCase(); // Col B
    const password = row[13] || ""; // Col N
    const estado = (row[15] || "").toString().toUpperCase(); // Col P
    
    // ✅ Para verificación, solo comparar username
    const isMatch = obj.checkOnly ? 
      username === obj.username.toString().toLowerCase() : 
      username + password === id;
    
    if (isMatch) {
      // Si es solo verificación de estado
      if (obj.checkOnly) {
        return {
          blocked: estado === "NO",  // true solo si está explícitamente bloqueado
          accesos: row[16] || "",    // Col Q
          status: estado
        };
      }
      
      // Si el usuario está bloqueado, no permitir login
      if (estado === "NO") {
        return { blocked: true };
      }
      
      // Login exitoso - registrar en LOG
      getSpreadsheetPersonal().getSheetByName('Log').appendRow([new Date(), row[2]]);
      
      return {
        success: true,
        data: [
          row[0],  // ID
          row[1],  // Usuario
          row[2],  // Nombre
          row[6],  // Cargo
          row[4],  // Empresa
          row[12], // Email
          row[13], // Contraseña
          row[14], // Foto
          row[16], // Accesos
          row[17]  // Firma
        ]
      };
    }
  }
  
  return obj.checkOnly ? 
    { blocked: true, status: "NOT_FOUND" } : 
    { success: false };
}

function actualizarUsuariologin(datos) {
  const hoja = getSpreadsheetPersonal().getSheetByName('PERSONAL');
  const lastRow = hoja.getLastRow();
  const folder = DriveApp.getFolderById(folderIdFirma);
  let urlFirma = "";

  // Leer solo la columna B (DNI/CE)
  const usuarios = hoja.getRange(2, 2, lastRow - 1, 1).getValues().flat();

  for (let i = 0; i < usuarios.length; i++) {
    const usuarioFila = String(usuarios[i]).toLowerCase();
    if (usuarioFila === datos.usuario.toLowerCase()) {
      const fila = i + 2;

      // Subir nueva firma si existe
      if (datos.firma && datos.firma.startsWith("data:image")) {
        const firmaBytes = Utilities.base64Decode(datos.firma.split(',')[1]);
        const blob = Utilities.newBlob(firmaBytes, 'image/png', `firma_${datos.usuario}.png`);
        const archivo = folder.createFile(blob);
        archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileId = archivo.getId();
        urlFirma = `https://lh5.googleusercontent.com/d/${fileId}`;
      } else {
        // Si no hay firma nueva, obtener la actual de la columna R (FIRMA, índice 18)
        urlFirma = hoja.getRange(fila, 18).getValue() || "";
      }

      // Actualizar datos en las columnas correspondientes
      hoja.getRange(fila, 2).setValue(datos.usuario);     // DNI/CE (columna B)
      hoja.getRange(fila, 3).setValue(datos.nombre);      // NOMBRES (columna C)
      hoja.getRange(fila, 5).setValue(datos.empresa);     // EMPRESA (columna E)
      hoja.getRange(fila, 7).setValue(datos.cargo);       // CARGO (columna G)
      hoja.getRange(fila, 13).setValue(datos.email);      // EMAIL (columna M)
      hoja.getRange(fila, 14).setValue(datos.password);   // CONTRASEÑA (columna N)
      hoja.getRange(fila, 15).setValue(datos.link);       // IMAGEN (columna O)
      hoja.getRange(fila, 17).setValue(datos.accesos);    // ACCESOS (columna Q)
      hoja.getRange(fila, 18).setValue(urlFirma);         // FIRMA (columna R)

      break;
    }
  }

  return urlFirma;
}

// Obtiene todos los registros
function getRecordsList() {
  const sheet = getSpreadsheetPersonal().getSheetByName('LISTAS');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { headersLista: [], data: [] };

  const headersLista = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return { headersLista, data };
}


function saveRecordsList(records) {
  const sheet = getSpreadsheetPersonal().getSheetByName('LISTAS');
  
  // Limpia las filas de datos anteriores
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  // Escribe los nuevos datos desde la fila 2
  sheet.getRange(2, 1, records.length, records[0].length).setValues(records);
  return 'Datos guardados';
}

// color de usuario
function getColor() {
  const sheet = getSpreadsheetPersonal().getSheetByName("RESUMEN");
  return sheet.getRange("J1").getValue(); // O cualquier celda donde guardes el color
}

function saveColor(color) {
  const sheet = getSpreadsheetPersonal().getSheetByName("RESUMEN");
  sheet.getRange("J1").setValue(color); // Guarda el color
}

//USUARIOS.
const HOJA = "PERSONAL";

function obtenerUsuariosPaginado(offset, limit, filtro = "") {
  const hoja = getSpreadsheetPersonal().getSheetByName(HOJA);
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) return { headers3: [], filas: [], total: 0 };

  const rango = hoja.getRange(1, 1, ultimaFila, 19).getValues();
  const columnas = [0, 1, 2, 6, 4, 5, 12, 13, 14, 15, 16];

  const headers3 = columnas.map(i => rango[0][i]);

  let filas = rango.slice(1)
    // ✅ Filtrar filas donde la columna B (índice 1) esté vacía
    .filter(fila => String(fila[1]).trim() !== "")
    // ✅ Luego seleccionamos solo las columnas necesarias
    .map(fila => columnas.map(i => fila[i]));

  filas = filas.reverse(); // Mostrar últimos primero

  if (filtro) {
    const texto = filtro.toLowerCase();
    filas = filas.filter(fila =>
      fila.some(celda => String(celda).toLowerCase().includes(texto))
    );
  }

  const paginados = filas.slice(offset, offset + limit);
  return { headers3, filas: paginados, total: filas.length };
}

//SIN ENVIAR EMAIL
// function agregarUsuario(data) {
//   const hoja = getSpreadsheetPersonal().getSheetByName(HOJA);
//   hoja.appendRow(data);
// }

//ENVIA EMAIL AL USUARIO
function agregarUsuario(data) {
  const hoja = getSpreadsheetPersonal().getSheetByName(HOJA);

  if (!data[0]) {
    const timestamp = Date.now().toString().slice(-7);
    data[0] = "E" + timestamp;
  }

  // Columnas destino en la hoja: [A, B, C, G, E, F, M, N, O, P, Q]
  const columnas = [1, 2, 3, 7, 5, 6, 13, 14, 15, 16, 17];
  const nuevaFila = hoja.getLastRow() + 1;

  columnas.forEach((col, i) => {
    hoja.getRange(nuevaFila, col).setValue(data[i]);
  });

  const usuario = data[1];
  const nombre = data[2];
  const email = data[13];
  const password = data[14];

  // ✅ NOTIFICACIÓN TELEGRAM
  try {
    notificarNuevoUsuario({
      nombre: nombre,
      usuario: usuario,
      empresa: empresa,
      cargo: cargo,
      email: email
    });
  } catch (error) {
    Logger.log('Error notificación Telegram: ' + error);
  }

  if (email) {
    const asunto = "Bienvenido a la plataforma";

    const cuerpoHtml = `
      <div style="font-family: Arial, sans-serif; background: #f4f4f4; padding: 20px; border-radius: 10px; max-width: 600px; margin: auto; border: 1px solid #ddd;">
        <h2 style="color: #2c3e50;">👋 ¡Hola, ${nombre}!</h2>
        <p style="font-size: 15px; color: #555;">Tu usuario ha sido registrado con éxito en nuestra plataforma.</p>
        <div style="background: #fff; padding: 15px 20px; border-radius: 8px; border: 1px solid #ccc; margin-top: 15px;">
          <p style="margin: 8px 0;"><strong>👤 Usuario:</strong> ${usuario}</p>
          <p style="margin: 8px 0;"><strong>🔐 Contraseña:</strong> ${password}</p>
        </div>
        <p style="margin-top: 20px; color: #333;">Puedes iniciar sesión con estos datos.</p>

        <div style="text-align: center; margin: 30px 0;">
          <a href="https://www.iassoma.com/ccl" target="_blank" 
             style="background-color: #2c3e50; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold;">
            🔗 Ingresar a la Plataforma
          </a>
        </div>

        <p style="font-size: 13px; color: #888;">No compartas esta información con terceros.</p>
        <hr style="margin-top: 30px;">
        <p style="text-align: center; font-size: 12px; color: #aaa;">
          © ${new Date().getFullYear()} BIOX-SIG - Todos los derechos reservados.
        </p>
      </div>
    `;

    // Enviar el correo con formato HTML
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      htmlBody: cuerpoHtml
    });
  }
}

function actualizarUsuario(data) {
  const hoja = getSpreadsheetPersonal().getSheetByName(HOJA);
  const id = String(data[0]).trim();
  const lastRow = hoja.getLastRow();
  const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
  const columnas = [1, 2, 3, 7, 5, 6, 13, 14, 15, 16, 17];

  for (let i = 0; i < ids.length; i++) {
    const idFila = String(ids[i][0]).trim();
    if (idFila === id) {
      const row = i + 2;
      columnas.forEach((col, j) => {
        hoja.getRange(row, col).setValue(data[j]);
      });
      
      // ✅ NOTIFICACIÓN TELEGRAM
      try {
        notificarUsuarioActualizado({
          usuario: data[1],
          modificadoPor: Session.getActiveUser().getEmail()
        });
      } catch (error) {
        Logger.log('Error notificación Telegram: ' + error);
      }
      
      return;
    }
  }
}


function eliminarUsuarioPorUsuario(usuario) {
  const hoja = getSpreadsheetPersonal().getSheetByName(HOJA);
  const lastRow = hoja.getLastRow();
  const usuarios = hoja.getRange(2, 2, lastRow - 1, 1).getValues(); 
  const usuarioBuscado = String(usuario).trim();

  for (let i = 0; i < usuarios.length; i++) {
    const usuarioFila = String(usuarios[i][0]).trim();
    if (usuarioFila === usuarioBuscado) {
      hoja.deleteRow(i + 2);
      
      // ✅ NOTIFICACIÓN TELEGRAM
      try {
        notificarUsuarioEliminado(usuario);
      } catch (error) {
        Logger.log('Error notificación Telegram: ' + error);
      }
      
      return;
    }
  }
}


function buscarDatosPorNumero(numero) {
  const hoja = getSpreadsheetPersonal().getSheetByName("PERSONAL");
  const datos = hoja.getRange("B2:M").getValues(); // Asegúrate de incluir hasta Q

  for (let i = 0; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(numero).trim()) {
      return {
        texto: datos[i][1],    // Columna C Nombre
        select1: datos[i][5],  // Columna G Cargo
        select2: datos[i][3],  // Columna E Empresa
        emailInput: datos[i][11] // Columna M Email
      };
    }
  }
  return null;
}

 function getTodasLasListas() {
  const cache = CacheService.getScriptCache();
  const cacheKey = "listas_globales";
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const libro = getSpreadsheetPersonal();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 📌 Obtener datos de MOF (columna C)
  const hojaMOF = libro.getSheetByName("MOF");
  const cargos = [...new Set(
    hojaMOF.getRange("C2:C" + hojaMOF.getLastRow()).getValues().map(r => r[0]).filter(Boolean)
  )];

  // 📌 Obtener datos de PERSONAL (columna C)
  const hojaTrabajadores = libro.getSheetByName("PERSONAL");
  const trabajadores = [...new Set(
    hojaTrabajadores.getRange("C2:C" + hojaTrabajadores.getLastRow()).getValues().map(r => r[0]).filter(Boolean)
  )];

  // 📌 Obtener datos de ACCESOS (columna B)
  const hojaAccesos = ss.getSheetByName("Accesos");
  const accesos = [...new Set(
    hojaAccesos.getRange("B2:B" + hojaAccesos.getLastRow()).getValues().map(r => r[0]).filter(Boolean)
  )];

  const hojaListas = libro.getSheetByName("LISTAS");
  const dataListas = hojaListas.getRange("B2:N" + hojaListas.getLastRow()).getValues();

  // Extraer cada columna individualmente
  const extractUnique = index =>
    [...new Set(dataListas.map(row => row[index]).filter(Boolean))];

  const listas = {
    // MOF, PERSONAL, ACCESOS
    cargos,
    trabajadores,
    accesos,
    areas: extractUnique(0),          // Col B
    inspectores: extractUnique(1),    // Col C
    lugares: extractUnique(2),        // Col D
    empresas: extractUnique(3),       // Col E
    estados: extractUnique(4),        // Col F
    gestiones: extractUnique(5),      // Col G
    desvios: extractUnique(6),        // Col H
    potenciales: extractUnique(7),    // Col I
    clasificaciones: extractUnique(8),// Col J
    riesgos: extractUnique(9),        // Col K
    capacitaciones: extractUnique(10),// Col L
    equipos: extractUnique(11),       // Col M
    procesos: extractUnique(12)       // Col N
  };

  // 🧠 Guardar en caché por 5 minutos (300 segundos)
  cache.put(cacheKey, JSON.stringify(listas), 300);
  return listas;
}


function enviarTelegram(mensaje) {
  const TOKEN = '8316348321:AAHyx9OczZdtoNuYi8OzPXx868c1tzhhwmc';
  const CHAT_ID = '6725665354'; // Reemplaza con tu ID numérico
  const URL = "https://api.telegram.org/bot" + TOKEN + "/sendMessage";
  
  const payload = {
    "chat_id": CHAT_ID,
    "text": mensaje,
    "parse_mode": "HTML"
  };
  
  const opciones = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    UrlFetchApp.fetch(URL, opciones);
  } catch (e) {
    console.error("Error enviando a Telegram: " + e);
  }
}

// ============================================
// FUNCIÓN: getIncompatibilidadData
// Descripción: Obtiene datos de incompatibilidades del sistema
// Ubicación sugerida: Agregar al archivo 3_CheckCode_gs.txt o 11_VerificacionesCode-gs.txt
// ============================================

/**
 * Obtiene datos de incompatibilidades con paginación y filtros
 * @param {number} offset - Posición inicial (para paginación)
 * @param {number} limit - Cantidad de registros a retornar
 * @param {string} search1 - Primer término de búsqueda (opcional)
 * @param {string} search2 - Segundo término de búsqueda (opcional)
 * @param {string} columnaFiltro1 - Columna para filtrar búsqueda 1 (opcional)
 * @param {string} columnaFiltro2 - Columna para filtrar búsqueda 2 (opcional)
 * @returns {Object} Objeto con headers, data y total
 */
function getIncompatibilidadData(offset, limit, search1, search2, columnaFiltro1, columnaFiltro2) {
  try {
    // OPCIÓN 1: Si tienes una hoja llamada "INCOMPATIBILIDADES" en el spreadsheet de Check
    const hoja = getCheckSpreadsheet().getSheetByName("INCOMPATIBILIDADES");
    
    // OPCIÓN 2: Si está en otro spreadsheet, descomenta y ajusta:
    // const hoja = SpreadsheetApp.openById("TU_SPREADSHEET_ID").getSheetByName("INCOMPATIBILIDADES");
    
    // Verificar si la hoja existe
    if (!hoja) {
      Logger.log("⚠️ Hoja INCOMPATIBILIDADES no encontrada");
      return {
        headers: [],
        data: [],
        total: 0,
        error: "Hoja no encontrada"
      };
    }

    const lastRow = hoja.getLastRow();
    const lastCol = hoja.getLastColumn();
    
    // Si no hay datos
    if (lastRow < 2) {
      return {
        headers: hoja.getRange(1, 1, 1, lastCol).getValues()[0],
        data: [],
        total: 0
      };
    }

    // Leer todos los datos (encabezados + registros)
    const allData = hoja.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headers = allData[0];
    let registros = allData.slice(1);

    // ============================================
    // APLICAR FILTROS DE BÚSQUEDA
    // ============================================
    const lowerSearch1 = (search1 || "").toLowerCase();
    const lowerSearch2 = (search2 || "").toLowerCase();

    if (lowerSearch1 || lowerSearch2) {
      registros = registros.filter(fila => {
        let pasaFiltro1 = true;
        let pasaFiltro2 = true;

        // Filtro 1
        if (lowerSearch1) {
          if (columnaFiltro1 && columnaFiltro1 !== "todos") {
            const colIndex = headers.indexOf(columnaFiltro1);
            if (colIndex !== -1) {
              pasaFiltro1 = fila[colIndex].toLowerCase().includes(lowerSearch1);
            }
          } else {
            pasaFiltro1 = fila.some(celda => 
              celda.toString().toLowerCase().includes(lowerSearch1)
            );
          }
        }

        // Filtro 2
        if (lowerSearch2) {
          if (columnaFiltro2 && columnaFiltro2 !== "todos") {
            const colIndex = headers.indexOf(columnaFiltro2);
            if (colIndex !== -1) {
              pasaFiltro2 = fila[colIndex].toLowerCase().includes(lowerSearch2);
            }
          } else {
            pasaFiltro2 = fila.some(celda => 
              celda.toString().toLowerCase().includes(lowerSearch2)
            );
          }
        }

        return pasaFiltro1 && pasaFiltro2;
      });
    }

    // ============================================
    // PAGINACIÓN (Mostrar los más recientes primero)
    // ============================================
    const totalFiltrados = registros.length;
    
    // Calcular rango para paginación inversa (últimos primero)
    const start = Math.max(totalFiltrados - offset - limit, 0);
    const end = totalFiltrados - offset;
    const paginados = registros.slice(start, end).reverse();

    return {
      headers: headers,
      data: paginados,
      total: totalFiltrados
    };

  } catch (error) {
    Logger.log("❌ Error en getIncompatibilidadData: " + error.message);
    Logger.log("Stack trace: " + error.stack);
    return {
      headers: [],
      data: [],
      total: 0,
      error: error.message
    };
  }
}

// ============================================
// FUNCIÓN SIMPLIFICADA (Si no necesitas filtros)
// ============================================

/**
 * Versión simplificada sin filtros - útil para pruebas iniciales
 * @returns {Object} Objeto con headers y data
 */
function getIncompatibilidadDataSimple() {
  try {
    const hoja = getCheckSpreadsheet().getSheetByName("INCOMPATIBILIDADES");
    
    if (!hoja) {
      return {
        headers: [],
        data: [],
        message: "Hoja INCOMPATIBILIDADES no encontrada"
      };
    }

    const lastRow = hoja.getLastRow();
    const lastCol = hoja.getLastColumn();
    
    if (lastRow < 2) {
      return {
        headers: hoja.getRange(1, 1, 1, lastCol).getValues()[0],
        data: [],
        message: "Sin datos"
      };
    }

    const allData = hoja.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    
    return {
      headers: allData[0],
      data: allData.slice(1)
    };

  } catch (error) {
    Logger.log("Error: " + error.message);
    return {
      headers: [],
      data: [],
      error: error.message
    };
  }
}

// ============================================
// FUNCIÓN DE PRUEBA
// ============================================

/**
 * Función para probar que getIncompatibilidadData funciona correctamente
 * Ejecuta esta función desde el editor de Apps Script para verificar
 */
function testGetIncompatibilidadData() {
  Logger.log("🔍 Iniciando prueba de getIncompatibilidadData...");
  
  // Prueba 1: Sin parámetros
  const resultado1 = getIncompatibilidadData(0, 10);
  Logger.log("📊 Prueba 1 - Sin filtros:");
  Logger.log("   Headers: " + resultado1.headers.length + " columnas");
  Logger.log("   Data: " + resultado1.data.length + " filas");
  Logger.log("   Total: " + resultado1.total + " registros");
  
  // Prueba 2: Con búsqueda
  const resultado2 = getIncompatibilidadData(0, 10, "test", "", "", "");
  Logger.log("📊 Prueba 2 - Con filtro 'test':");
  Logger.log("   Data: " + resultado2.data.length + " filas encontradas");
  
  // Prueba 3: Versión simple
  const resultado3 = getIncompatibilidadDataSimple();
  Logger.log("📊 Prueba 3 - Versión simple:");
  Logger.log("   Headers: " + resultado3.headers.length);
  Logger.log("   Data: " + resultado3.data.length);
  
  Logger.log("✅ Prueba completada");
}

// ============================================
// FUNCIONES AUXILIARES OPCIONALES
// ============================================

/**
 * Obtener encabezados de la tabla de incompatibilidades
 * Útil para poblar selectores de columnas en el frontend
 */
function getIncompatibilidadHeaders() {
  try {
    const hoja = getCheckSpreadsheet().getSheetByName("INCOMPATIBILIDADES");
    if (!hoja) return [];
    
    const lastCol = hoja.getLastColumn();
    return hoja.getRange(1, 1, 1, lastCol).getValues()[0];
  } catch (error) {
    Logger.log("Error obteniendo headers: " + error.message);
    return [];
  }
}

/**
 * Agregar un nuevo registro de incompatibilidad
 * @param {Array} data - Array con los datos a agregar
 */
function agregarIncompatibilidad(data) {
  try {
    const hoja = getCheckSpreadsheet().getSheetByName("INCOMPATIBILIDADES");
    if (!hoja) {
      throw new Error("Hoja INCOMPATIBILIDADES no encontrada");
    }
    
    // Agregar timestamp si no viene en los datos
    const timestamp = new Date();
    const dataConFecha = [...data, timestamp];
    
    hoja.appendRow(dataConFecha);
    return { success: true, message: "Registro agregado correctamente" };
    
  } catch (error) {
    Logger.log("Error agregando incompatibilidad: " + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Actualizar un registro de incompatibilidad por ID
 * @param {string} id - ID del registro a actualizar
 * @param {Array} data - Nuevos datos
 */
function actualizarIncompatibilidad(id, data) {
  try {
    const hoja = getCheckSpreadsheet().getSheetByName("INCOMPATIBILIDADES");
    if (!hoja) {
      throw new Error("Hoja INCOMPATIBILIDADES no encontrada");
    }
    
    const lastRow = hoja.getLastRow();
    const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0].toString() === id.toString()) {
        const fila = i + 2;
        hoja.getRange(fila, 1, 1, data.length).setValues([data]);
        return { success: true, message: "Registro actualizado" };
      }
    }
    
    return { success: false, message: "ID no encontrado" };
    
  } catch (error) {
    Logger.log("Error actualizando incompatibilidad: " + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Eliminar un registro de incompatibilidad por ID
 * @param {string} id - ID del registro a eliminar
 */
function eliminarIncompatibilidad(id) {
  try {
    const hoja = getCheckSpreadsheet().getSheetByName("INCOMPATIBILIDADES");
    if (!hoja) {
      throw new Error("Hoja INCOMPATIBILIDADES no encontrada");
    }
    
    const lastRow = hoja.getLastRow();
    const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0].toString() === id.toString()) {
        const fila = i + 2;
        hoja.deleteRow(fila);
        return { success: true, message: "Registro eliminado" };
      }
    }
    
    return { success: false, message: "ID no encontrado" };
    
  } catch (error) {
    Logger.log("Error eliminando incompatibilidad: " + error.message);
    return { success: false, error: error.message };
  }
}