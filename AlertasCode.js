/**
 * Obtiene alertas de vencimiento estructuradas para mostrar en una tabla.
 * Solo muestra la ÚLTIMA entrega de cada combinación DNI + Producto + Variante.
 * Si el trabajador ya recibió una entrega nueva, las alertas anteriores se cierran.
 * @param {string} dniLogin - El DNI del usuario logueado.
 */
function obtenerAlertasVencimientos(dniLogin) {
  try {
    const ss = getSpreadsheetEPP();
    const shReg = ss.getSheetByName(SHEPP.REGISTRO);
    const correoActual = Session.getActiveUser().getEmail();

    const ADMIN_EMAIL = "tu_correo_admin@gmail.com";
    const esAdmin = (correoActual === ADMIN_EMAIL || !dniLogin);

    const data = shReg.getDataRange().getValues();
    const hoy = new Date();

    const COL_DNI = IDX.REG.DNI - 1;
    const COL_PRODUCTO = IDX.REG.PRODUCTO - 1;
    const COL_VARIANTE = IDX.REG.VARIANTE - 1;
    const COL_VENC = IDX.REG.FECHA_VENCIMIENTO - 1;
    const COL_OP = IDX.REG.OPERACION - 1;
    const COL_NOMBRES = IDX.REG.NOMBRES - 1;
    const COL_FECHA = IDX.REG.FECHA - 1;

    // PASO 1: Agrupar entregas y quedarse solo con la MÁS RECIENTE
    // por cada combinación DNI + Producto + Variante
    const ultimasPorItem = {};

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];

      if (!esAdmin && _str(fila[COL_DNI]) !== _str(dniLogin)) continue;
      if (_str(fila[COL_OP]) !== 'Entrega') continue;

      const dni = _str(fila[COL_DNI]);
      const producto = _str(fila[COL_PRODUCTO]);
      const variante = _str(fila[COL_VARIANTE] || '');
      const clave = dni + '|' + producto + '|' + variante;

      const fechaEntrega = new Date(fila[COL_FECHA]);
      if (isNaN(fechaEntrega.getTime())) continue;

      // Solo guardar la más reciente por cada clave
      if (!ultimasPorItem[clave] || fechaEntrega > ultimasPorItem[clave].fechaEntrega) {
        ultimasPorItem[clave] = { fila, fechaEntrega };
      }
    }

    // PASO 2: Generar alertas solo de las últimas entregas
    const alertas = [];

    for (const clave in ultimasPorItem) {
      const { fila } = ultimasPorItem[clave];

      const fechaVencRaw = fila[COL_VENC];
      if (!fechaVencRaw) continue;

      const fechaVenc = new Date(fechaVencRaw);
      if (isNaN(fechaVenc.getTime())) continue;

      const diffDias = Math.ceil((fechaVenc - hoy) / (1000 * 60 * 60 * 24));
      const fechaFormateada = Utilities.formatDate(fechaVenc, "GMT-5", "dd/MM/yyyy");

      if (diffDias <= 0) {
        alertas.push({
          producto: _str(fila[COL_PRODUCTO]),
          trabajador: _str(fila[COL_NOMBRES]),
          fecha: fechaFormateada,
          estado: "VENCIDO",
          clase: "fila-vencida",
          badge: "bg-rojo"
        });
      }
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
