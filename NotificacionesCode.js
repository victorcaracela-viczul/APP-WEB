// ============================================================
//  NotificacionesCode.js — Push Notifications vía Cloudflare Worker
//  Envía notificaciones push a dispositivos de trabajadores
// ============================================================

// URL de tu Cloudflare Worker (ACTUALIZAR con tu dominio real)
const PUSH_WORKER_URL = 'https://viczul.com';

// Token secreto para autenticar llamadas GAS → Worker
// IMPORTANTE: Configurar el mismo valor como variable de entorno
// PUSH_AUTH_TOKEN en tu Cloudflare Worker
const PUSH_AUTH_TOKEN = 'adecco_push_2026_secret_token_xyz123';

/**
 * Enviar push a UN trabajador por DNI
 */
function enviarPushNotification(dni, title, body, tag) {
  try {
    if (!dni || !title) return { ok: false, error: 'dni y title requeridos' };

    const payload = {
      token: PUSH_AUTH_TOKEN,
      dni: String(dni).trim(),
      title: title,
      body: body || '',
      tag: tag || 'general',
      url: '/'
    };

    const resp = UrlFetchApp.fetch(PUSH_WORKER_URL + '/api/push/send', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(resp.getContentText());
    Logger.log('Push enviado a ' + dni + ': ' + JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log('Error enviando push: ' + e.message);
    return { ok: false, error: e.message };
  }
}

/**
 * Enviar push a VARIOS trabajadores por DNI[]
 */
function enviarPushBulk(dnis, title, body, tag) {
  try {
    if (!Array.isArray(dnis) || !dnis.length || !title) {
      return { ok: false, error: 'dnis[] y title requeridos' };
    }

    const payload = {
      token: PUSH_AUTH_TOKEN,
      dnis: dnis.map(d => String(d).trim()),
      title: title,
      body: body || '',
      tag: tag || 'general',
      url: '/'
    };

    const resp = UrlFetchApp.fetch(PUSH_WORKER_URL + '/api/push/send-bulk', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(resp.getContentText());
    Logger.log('Push bulk enviado a ' + dnis.length + ' usuarios: ' + JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log('Error enviando push bulk: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ============================================================
//  FUNCIONES DE NOTIFICACIÓN POR MÓDULO (automáticas)
// ============================================================

function notificarEntregaEPP(dni, producto, variante) {
  const desc = producto + (variante ? ' (' + variante + ')' : '');
  return enviarPushNotification(
    dni,
    'EPP Asignado',
    'Se te asignó: ' + desc + '. Ingresa para firmar la recepción.',
    'epp-entrega'
  );
}

function notificarConfirmacionEPP(dniSupervisor, trabajadorNombre, producto, accion) {
  const titulo = accion === 'confirmado' ? 'EPP Confirmado' : 'EPP Rechazado';
  const cuerpo = trabajadorNombre + ' ' + accion + ' la recepción de: ' + producto;
  return enviarPushNotification(dniSupervisor, titulo, cuerpo, 'epp-confirmacion');
}

function notificarCapacitacion(dnis, tema, fecha) {
  const body = fecha
    ? 'Capacitación: ' + tema + ' programada para ' + fecha + '. Revisa tu app.'
    : 'Tienes una capacitación asignada: ' + tema + '. Revisa tu app.';
  return enviarPushBulk(dnis, 'Capacitación', body, 'capacitacion');
}

function enviarNotificacionManual(dni, titulo, mensaje) {
  return enviarPushNotification(dni, titulo, mensaje, 'manual');
}

function notificarATodos(titulo, mensaje) {
  try {
    const hoja = getSpreadsheetPersonal().getSheetByName('PERSONAL');
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return { ok: false, error: 'No hay trabajadores' };

    const data = hoja.getRange(2, 1, lastRow - 1, 17).getValues();
    const dnis = [];

    for (let i = 0; i < data.length; i++) {
      const estado = (data[i][15] || '').toString().toUpperCase(); // Col P = estado
      const dni = (data[i][1] || '').toString().trim(); // Col B = DNI
      if (dni && estado !== 'NO') {
        dnis.push(dni);
      }
    }

    if (!dnis.length) return { ok: false, error: 'No hay trabajadores activos' };
    return enviarPushBulk(dnis, titulo, mensaje, 'general');
  } catch (e) {
    Logger.log('Error notificando a todos: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ============================================================
//  FUNCIONES MANUALES — Llamadas desde el frontend (admin)
//  Usadas por google.script.run desde EPPMaestro y Capacitaciones
// ============================================================

/**
 * Enviar alerta manual de EPP desde el panel de administración
 * @param {string} modo - "individual" o "todos"
 * @param {string} [dni] - DNI del trabajador (solo si modo="individual")
 * @param {string} titulo - Título de la alerta
 * @param {string} mensaje - Mensaje de la alerta
 * @returns {Object} { ok, sent, ... }
 */
function enviarAlertaEPP(modo, dni, titulo, mensaje) {
  try {
    if (modo === 'todos') {
      return notificarATodos(
        titulo || 'Alerta EPP',
        mensaje || 'Revisa tu módulo de EPP. Tienes actualizaciones pendientes.'
      );
    } else {
      if (!dni) return { ok: false, error: 'DNI requerido para notificación individual' };
      return enviarPushNotification(
        dni,
        titulo || 'Alerta EPP',
        mensaje || 'Revisa tu módulo de EPP. Tienes actualizaciones pendientes.',
        'epp-manual'
      );
    }
  } catch (e) {
    Logger.log('Error en alerta EPP manual: ' + e.message);
    return { ok: false, error: e.message };
  }
}

/**
 * Enviar alerta manual de Capacitaciones desde el panel de administración
 * @param {string} modo - "individual", "seleccion" o "todos"
 * @param {string|string[]} dniOrDnis - DNI o array de DNIs
 * @param {string} titulo - Título de la alerta
 * @param {string} mensaje - Mensaje de la alerta
 * @returns {Object} { ok, sent, ... }
 */
function enviarAlertaCapacitacion(modo, dniOrDnis, titulo, mensaje) {
  try {
    const tit = titulo || 'Alerta Capacitación';
    const msg = mensaje || 'Tienes una capacitación pendiente. Revisa tu app.';

    if (modo === 'todos') {
      return notificarATodos(tit, msg);
    } else if (modo === 'seleccion' && Array.isArray(dniOrDnis)) {
      return enviarPushBulk(dniOrDnis, tit, msg, 'cap-manual');
    } else {
      if (!dniOrDnis) return { ok: false, error: 'DNI requerido' };
      return enviarPushNotification(String(dniOrDnis), tit, msg, 'cap-manual');
    }
  } catch (e) {
    Logger.log('Error en alerta Cap manual: ' + e.message);
    return { ok: false, error: e.message };
  }
}

/**
 * Obtener lista de trabajadores activos para los selects de notificación
 * Devuelve [{dni, nombre, cargo, empresa}]
 */
function obtenerTrabajadoresParaNotificar() {
  try {
    const hoja = getSpreadsheetPersonal().getSheetByName('PERSONAL');
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return [];

    const data = hoja.getRange(2, 1, lastRow - 1, 17).getValues();
    const trabajadores = [];

    for (let i = 0; i < data.length; i++) {
      const estado = (data[i][15] || '').toString().toUpperCase(); // Col P = estado
      const dni = (data[i][1] || '').toString().trim(); // Col B = DNI
      const nombre = (data[i][2] || '').toString().trim(); // Col C = Nombre
      const cargo = (data[i][3] || '').toString().trim(); // Col D = Cargo
      const empresa = (data[i][4] || '').toString().trim(); // Col E = Empresa
      if (dni && estado !== 'NO') {
        trabajadores.push({ dni: dni, nombre: nombre, cargo: cargo, empresa: empresa });
      }
    }

    return trabajadores.sort(function(a, b) {
      return (a.nombre || '').localeCompare(b.nombre || '');
    });
  } catch (e) {
    Logger.log('Error obteniendo trabajadores: ' + e.message);
    return [];
  }
}

// ============================================================
//  TEST — Ejecutar desde el editor GAS para diagnosticar conexión
//  En el editor: seleccionar esta función y clic en ▶ Run
// ============================================================
function testPushConnection() {
  // Test 1: Verificar que el Worker responde (GET sin auth)
  Logger.log('=== TEST 1: Verificar Worker ===');
  try {
    var resp1 = UrlFetchApp.fetch(PUSH_WORKER_URL + '/api/push/test', {
      muteHttpExceptions: true
    });
    Logger.log('Status: ' + resp1.getResponseCode());
    Logger.log('Body: ' + resp1.getContentText());
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
  }

  // Test 2: Enviar push de prueba (POST con auth)
  Logger.log('=== TEST 2: Enviar push con auth ===');
  Logger.log('URL: ' + PUSH_WORKER_URL + '/api/push/send');
  Logger.log('Token que envío (longitud): ' + PUSH_AUTH_TOKEN.length);
  Logger.log('Token que envío (primeros 10): ' + PUSH_AUTH_TOKEN.substring(0, 10));
  try {
    var payload = {
      token: PUSH_AUTH_TOKEN,
      dni: '99999999',
      title: 'Test de conexión',
      body: 'Si ves esto, la conexión funciona',
      tag: 'test'
    };
    Logger.log('Payload: ' + JSON.stringify(payload));
    var resp2 = UrlFetchApp.fetch(PUSH_WORKER_URL + '/api/push/send', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log('Status: ' + resp2.getResponseCode());
    Logger.log('Body: ' + resp2.getContentText());
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
  }

  Logger.log('=== FIN DE TESTS ===');
}
