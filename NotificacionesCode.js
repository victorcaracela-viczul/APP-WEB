// ============================================================
//  NotificacionesCode.js — Push Notifications vía Cloudflare Worker
//  Envía notificaciones push a dispositivos de trabajadores
// ============================================================

// URL de tu Cloudflare Worker (cambiar si tu dominio es diferente)
const PUSH_WORKER_URL = 'https://sistema-gestion.tu-dominio.workers.dev';

// Token secreto para autenticar llamadas GAS → Worker
// IMPORTANTE: Configurar el mismo valor como variable de entorno
// PUSH_AUTH_TOKEN en tu Cloudflare Worker
const PUSH_AUTH_TOKEN = 'adecco-isos-push-secret-2024';

/**
 * Enviar push a UN trabajador por DNI
 * @param {string} dni - DNI del trabajador
 * @param {string} title - Título de la notificación
 * @param {string} body - Cuerpo del mensaje
 * @param {string} [tag] - Tag para agrupar/reemplazar notificaciones
 * @returns {Object} { ok, sent }
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
 * @param {string[]} dnis - Array de DNIs
 * @param {string} title - Título de la notificación
 * @param {string} body - Cuerpo del mensaje
 * @param {string} [tag] - Tag para agrupar
 * @returns {Object} { ok, sent, dnis }
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
//  FUNCIONES DE NOTIFICACIÓN POR MÓDULO
// ============================================================

/**
 * Notificar al trabajador que tiene un EPP pendiente de firma
 * Se llama desde registrarEntrega() en EppCode.js
 */
function notificarEntregaEPP(dni, producto, variante) {
  const desc = producto + (variante ? ' (' + variante + ')' : '');
  return enviarPushNotification(
    dni,
    'EPP Asignado',
    'Se te asignó: ' + desc + '. Ingresa para firmar la recepción.',
    'epp-entrega'
  );
}

/**
 * Notificar al supervisor que el trabajador confirmó/rechazó
 * Se llama desde confirmarEntregaEpp() / rechazarEntregaEpp()
 */
function notificarConfirmacionEPP(dniSupervisor, trabajadorNombre, producto, accion) {
  const titulo = accion === 'confirmado' ? 'EPP Confirmado' : 'EPP Rechazado';
  const cuerpo = trabajadorNombre + ' ' + accion + ' la recepción de: ' + producto;
  return enviarPushNotification(dniSupervisor, titulo, cuerpo, 'epp-confirmacion');
}

/**
 * Notificar a trabajadores sobre una capacitación
 * @param {string[]} dnis - DNIs de los trabajadores convocados
 * @param {string} tema - Nombre de la capacitación
 * @param {string} [fecha] - Fecha de la capacitación
 */
function notificarCapacitacion(dnis, tema, fecha) {
  const body = fecha
    ? 'Capacitación: ' + tema + ' programada para ' + fecha + '. Revisa tu app.'
    : 'Tienes una capacitación asignada: ' + tema + '. Revisa tu app.';
  return enviarPushBulk(dnis, 'Capacitación', body, 'capacitacion');
}

/**
 * Notificación genérica - se puede llamar desde cualquier módulo
 * Expuesta como función global para usar desde el frontend si se necesita
 */
function enviarNotificacionManual(dni, titulo, mensaje) {
  return enviarPushNotification(dni, titulo, mensaje, 'manual');
}

/**
 * Enviar notificación a todos los trabajadores activos
 * Útil para avisos generales
 */
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
