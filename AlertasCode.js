/**
 * Obtiene alertas de vencimiento estructuradas para mostrar en una tabla.
 * @param {string} dniLogin - El DNI del usuario logueado.
 */
function obtenerAlertasVencimientos(dniLogin) {
  try {
    const ss = getSpreadsheetEPP();
    const shReg = ss.getSheetByName(SHEPP.REGISTRO);
    const correoActual = Session.getActiveUser().getEmail();
    
    // 1. DEFINIR SI ES ADMIN (Cambia el correo por el tuyo)
    const ADMIN_EMAIL = "tu_correo_admin@gmail.com"; 
    const esAdmin = (correoActual === ADMIN_EMAIL || !dniLogin);

    const data = shReg.getDataRange().getValues();
    const hoy = new Date();
    const alertas = [];

    // Índices basados en tu objeto IDX.REG
    const COL_DNI = IDX.REG.DNI - 1;               
    const COL_PRODUCTO = IDX.REG.PRODUCTO - 1;     
    const COL_VENC = IDX.REG.FECHA_VENCIMIENTO - 1;
    const COL_OP = IDX.REG.OPERACION - 1;          
    const COL_NOMBRES = IDX.REG.NOMBRES - 1;       

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];

      // Filtro de seguridad: Si no es admin, solo ve su DNI
      if (!esAdmin && _str(fila[COL_DNI]) !== _str(dniLogin)) continue;

      // Solo procesar registros de "Entrega"
      if (_str(fila[COL_OP]) !== 'Entrega') continue;

      const fechaVencRaw = fila[COL_VENC];
      if (!fechaVencRaw) continue;

      const fechaVenc = new Date(fechaVencRaw);
      if (isNaN(fechaVenc.getTime())) continue;

      // Cálculo de días restantes
      const diffDias = Math.ceil((fechaVenc - hoy) / (1000 * 60 * 60 * 24));

      // Formatear fecha para mostrar en la tabla (dd/mm/yyyy)
      const fechaFormateada = Utilities.formatDate(fechaVenc, "GMT-5", "dd/MM/yyyy");

      // CASO 1: YA VENCIDO (Rojo)
      if (diffDias <= 0) {
        alertas.push({
          producto: _str(fila[COL_PRODUCTO]),
          trabajador: _str(fila[COL_NOMBRES]), // Útil para la vista de Admin
          fecha: fechaFormateada,
          estado: "VENCIDO",
          clase: "fila-vencida", // Clase CSS para la fila
          badge: "bg-rojo"       // Clase CSS para el circulito/etiqueta
        });
      } 
      // CASO 2: POR VENCER (Naranja - 15 días de anticipación)
      else if (diffDias <= 15) {
        alertas.push({
          producto: _str(fila[COL_PRODUCTO]),
          trabajador: _str(fila[COL_NOMBRES]),
          fecha: fechaFormateada,
          estado: `VENCE EN ${diffDias} DÍAS`,
          clase: "fila-proxima",
          badge: "bg-naranja"
        });
      }
    }

    // Ordenar: Primero los vencidos (Rojo) y luego los próximos (Naranja)
    return alertas.sort((a, b) => (a.badge === 'bg-rojo' ? -1 : 1)).slice(0, 15);

  } catch (e) {
    console.error("Error en alertas: " + e.message);
    return [];
  }
}

/**
 * Verifica si un trabajador tiene EPPs pendientes de firma.
 * Usado para mostrar aviso al login.
 * @param {string} dniLogin - DNI del usuario logueado
 * @returns {Object} { tieneVencimientos, tienePendientes, totalPendientes }
 */
function verificarAlertasCompletas(dniLogin) {
  try {
    const alertasVenc = obtenerAlertasVencimientos(dniLogin);
    const pendientes = (typeof obtenerEntregasPendientes === 'function')
      ? obtenerEntregasPendientes(dniLogin)
      : [];
    return {
      tieneVencimientos: alertasVenc && alertasVenc.length > 0,
      tienePendientes: pendientes && pendientes.length > 0,
      totalPendientes: pendientes ? pendientes.length : 0
    };
  } catch (e) {
    console.error("Error en verificarAlertasCompletas: " + e.message);
    return { tieneVencimientos: false, tienePendientes: false, totalPendientes: 0 };
  }
}
