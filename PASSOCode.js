// =============================================
// PASSO - Programa Anual de SSO (Backend)
// =============================================

// Helper: dias en un mes
function _diasEnMes(mes, anio) {
  return new Date(anio, mes, 0).getDate();
}

// Helper: obtener o crear hoja REUNIONES
function _getHojaReuniones() {
  const ss = getSpreadsheetCapacitaciones();
  let hoja = ss.getSheetByName("REUNIONES");
  if (!hoja) {
    hoja = ss.insertSheet("REUNIONES");
    hoja.appendRow(["ID", "Tema", "Responsable", "Gerencia", "Area", "MesProgramado", "MesEjecutado", "Anio"]);
  }
  return hoja;
}

// =============================================
// 1. INSPECCIONES - Datos del programa anual
// =============================================
function obtenerDatosPASSOInspecciones() {
  try {
    const anio = new Date().getFullYear();
    const ssCheck = getCheckSpreadsheet();

    // Leer INVENTARIO
    const hojaInv = ssCheck.getSheetByName("INVENTARIO");
    const lastRowInv = hojaInv.getLastRow();
    if (lastRowInv < 3) return { actividades: [] };

    const invData = hojaInv.getRange(3, 1, lastRowInv - 2, 18).getValues();

    // Filtrar equipos activos con frecuencia > 0
    const equipos = [];
    for (let i = 0; i < invData.length; i++) {
      const estado = String(invData[i][14] || "").toLowerCase();
      if (estado === "retirado") continue;

      const frecDias = parseInt(invData[i][13]) || 0;
      if (frecDias === 0) continue;

      equipos.push({
        numero: invData[i][0],
        equipo: String(invData[i][3] || "").trim(),
        codigo: String(invData[i][4] || "").trim(),
        area: String(invData[i][7] || "").trim(),
        frecuenciaDias: frecDias
      });
    }

    // Leer B DATOS de check para contar ejecuciones
    const hojaBD = ssCheck.getSheetByName("B DATOS");
    const lastRowBD = hojaBD.getLastRow();
    const ejecucionesPorEquipoMes = {};

    if (lastRowBD > 1) {
      const bdData = hojaBD.getRange(2, 1, lastRowBD - 1, 10).getValues();

      for (let i = 0; i < bdData.length; i++) {
        const equipoNombre = String(bdData[i][2] || "").trim();
        const fecha = bdData[i][9];

        if (!equipoNombre || !(fecha instanceof Date)) continue;
        if (fecha.getFullYear() !== anio) continue;

        const mes = fecha.getMonth(); // 0-11
        const clave = equipoNombre + "_" + mes;
        ejecucionesPorEquipoMes[clave] = (ejecucionesPorEquipoMes[clave] || 0) + 1;
      }
    }

    // Construir actividades con P/E por mes
    const actividades = equipos.map(eq => {
      const meses = [];

      for (let m = 0; m < 12; m++) {
        const diasMes = _diasEnMes(m + 1, anio);
        let programado = 0;

        if (eq.frecuenciaDias === 1) {
          programado = diasMes;
        } else if (eq.frecuenciaDias <= 7) {
          programado = Math.floor(diasMes / eq.frecuenciaDias);
        } else if (eq.frecuenciaDias <= 15) {
          programado = Math.floor(diasMes / eq.frecuenciaDias);
        } else if (eq.frecuenciaDias <= 31) {
          programado = 1;
        } else if (eq.frecuenciaDias <= 92) {
          // Trimestral: solo en meses 0,3,6,9
          programado = (m % 3 === 0) ? 1 : 0;
        } else {
          // Semestral o mayor
          programado = (m % 6 === 0) ? 1 : 0;
        }

        const clave = eq.equipo + "_" + m;
        const ejecutado = ejecucionesPorEquipoMes[clave] || 0;

        meses.push({ programado, ejecutado });
      }

      return {
        equipo: eq.equipo,
        codigo: eq.codigo,
        area: eq.area,
        frecuenciaDias: eq.frecuenciaDias,
        meses
      };
    });

    return { actividades, anio };
  } catch (error) {
    Logger.log("Error en obtenerDatosPASSOInspecciones: " + error.message);
    return { actividades: [], error: error.message };
  }
}

// =============================================
// 2. CAPACITACIONES / ENTRENAMIENTOS
// =============================================
function obtenerDatosPASSOCapacitaciones(tipo) {
  try {
    const anio = new Date().getFullYear();
    const ss = getSpreadsheetCapacitaciones();
    const hojaMatriz = ss.getSheetByName("Matriz");
    const hojaBD = ss.getSheetByName("B DATOS");
    const lastCol = hojaMatriz.getLastColumn();
    const numCursos = lastCol - 4;

    if (numCursos <= 0) return { actividades: [], anio };

    // Leer filas de metadata
    const tiposProg = hojaMatriz.getRange(2, 5, 1, numCursos).getValues()[0];
    const responsables = hojaMatriz.getRange(3, 5, 1, numCursos).getValues()[0];
    const gerencias = hojaMatriz.getRange(4, 5, 1, numCursos).getValues()[0];
    const programaciones = hojaMatriz.getRange(6, 5, 1, numCursos).getValues()[0];
    const areas = hojaMatriz.getRange(12, 5, 1, numCursos).getValues()[0];
    const cursos = hojaMatriz.getRange(15, 5, 1, numCursos).getValues()[0];

    // Leer B DATOS para ejecuciones
    const lastRowBD = hojaBD.getLastRow();
    const ejecucionesPorCursoMes = {};

    if (lastRowBD > 1) {
      const bdData = hojaBD.getRange(2, 1, lastRowBD - 1, 10).getValues();

      for (let i = 0; i < bdData.length; i++) {
        const tema = String(bdData[i][4] || "").trim();
        const fecha = bdData[i][7];
        const estado = String(bdData[i][9] || "");

        if (!tema || !(fecha instanceof Date)) continue;
        if (fecha.getFullYear() !== anio) continue;
        if (estado !== "Aprobado") continue;

        const mes = fecha.getMonth();
        const clave = tema + "_" + mes;
        ejecucionesPorCursoMes[clave] = true;
      }
    }

    // Filtrar por tipo y construir actividades
    const actividades = [];

    for (let j = 0; j < numCursos; j++) {
      const tipoCol = String(tiposProg[j] || "").trim();
      if (tipoCol !== tipo) continue;

      const cursoNombre = String(cursos[j] || "").trim();
      if (!cursoNombre) continue;

      // Determinar mes programado
      let mesProgramado = -1;
      const prog = programaciones[j];

      if (prog instanceof Date && !isNaN(prog)) {
        if (prog.getFullYear() === anio) {
          mesProgramado = prog.getMonth();
        }
      } else if (typeof prog === "string" && prog.trim() !== "") {
        const match = prog.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (match) {
          const fechaProg = new Date(match[3] + "-" + match[2] + "-" + match[1]);
          if (fechaProg.getFullYear() === anio) {
            mesProgramado = fechaProg.getMonth();
          }
        }
      }

      const meses = [];
      for (let m = 0; m < 12; m++) {
        const programado = (m === mesProgramado);
        const clave = cursoNombre + "_" + m;
        const ejecutado = !!ejecucionesPorCursoMes[clave];

        meses.push({ programado, ejecutado });
      }

      actividades.push({
        curso: cursoNombre,
        responsable: String(responsables[j] || ""),
        gerencia: String(gerencias[j] || ""),
        area: String(areas[j] || ""),
        meses
      });
    }

    return { actividades, anio };
  } catch (error) {
    Logger.log("Error en obtenerDatosPASSOCapacitaciones: " + error.message);
    return { actividades: [], error: error.message };
  }
}

// =============================================
// 3. REUNIONES - CRUD
// =============================================
function obtenerReuniones() {
  try {
    const hoja = _getHojaReuniones();
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return { reuniones: [] };

    const datos = hoja.getRange(2, 1, lastRow - 1, 8).getValues();
    const reuniones = datos.map(r => ({
      id: r[0],
      tema: r[1],
      responsable: r[2],
      gerencia: r[3],
      area: r[4],
      mesProgramado: r[5],
      mesEjecutado: r[6],
      anio: r[7]
    }));

    return { reuniones };
  } catch (error) {
    Logger.log("Error en obtenerReuniones: " + error.message);
    return { reuniones: [], error: error.message };
  }
}

function agregarReunion(data) {
  try {
    const hoja = _getHojaReuniones();
    const id = "R" + Date.now();
    const fila = [
      id,
      data[0] || "",  // tema
      data[1] || "",  // responsable
      data[2] || "",  // gerencia
      data[3] || "",  // area
      data[4] || "",  // mesProgramado (1-12)
      "",             // mesEjecutado (vacío)
      data[5] || new Date().getFullYear() // anio
    ];
    hoja.appendRow(fila);
    return { success: true, id: id };
  } catch (error) {
    Logger.log("Error en agregarReunion: " + error.message);
    return { success: false, error: error.message };
  }
}

function actualizarReunion(data) {
  try {
    const hoja = _getHojaReuniones();
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return { success: false, error: "No hay reuniones" };

    const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
    const idBuscado = String(data[0]).trim();

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === idBuscado) {
        const fila = [
          data[0],  // id
          data[1],  // tema
          data[2],  // responsable
          data[3],  // gerencia
          data[4],  // area
          data[5],  // mesProgramado
          data[6],  // mesEjecutado
          data[7]   // anio
        ];
        hoja.getRange(i + 2, 1, 1, 8).setValues([fila]);
        return { success: true };
      }
    }
    return { success: false, error: "Reunión no encontrada" };
  } catch (error) {
    Logger.log("Error en actualizarReunion: " + error.message);
    return { success: false, error: error.message };
  }
}

function eliminarReunion(id) {
  try {
    const hoja = _getHojaReuniones();
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return { success: false, error: "No hay reuniones" };

    const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
    const idBuscado = String(id).trim();

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === idBuscado) {
        hoja.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, error: "Reunión no encontrada" };
  } catch (error) {
    Logger.log("Error en eliminarReunion: " + error.message);
    return { success: false, error: error.message };
  }
}

function marcarReunionEjecutada(id, mesEjecutado) {
  try {
    const hoja = _getHojaReuniones();
    const lastRow = hoja.getLastRow();
    if (lastRow < 2) return { success: false, error: "No hay reuniones" };

    const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
    const idBuscado = String(id).trim();

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === idBuscado) {
        hoja.getRange(i + 2, 7).setValue(mesEjecutado); // Columna G = MesEjecutado
        return { success: true };
      }
    }
    return { success: false, error: "Reunión no encontrada" };
  } catch (error) {
    Logger.log("Error en marcarReunionEjecutada: " + error.message);
    return { success: false, error: error.message };
  }
}

// =============================================
// 4. DATOS COMBINADOS PASSO (todas las tabs)
// =============================================
function obtenerDatosPASSOCompleto() {
  try {
    const inspecciones = obtenerDatosPASSOInspecciones();
    const capacitaciones = obtenerDatosPASSOCapacitaciones("Capacitación");
    const reunionesData = obtenerReuniones();
    const entrenamientos = obtenerDatosPASSOCapacitaciones("Entrenamiento");

    // Convertir reuniones a formato calendario
    const anio = new Date().getFullYear();
    const reunionesActividades = reunionesData.reuniones
      .filter(r => Number(r.anio) === anio)
      .map(r => {
        const meses = [];
        for (let m = 0; m < 12; m++) {
          meses.push({
            programado: Number(r.mesProgramado) === (m + 1),
            ejecutado: Number(r.mesEjecutado) === (m + 1)
          });
        }
        return {
          id: r.id,
          curso: r.tema,
          responsable: r.responsable,
          gerencia: r.gerencia,
          area: r.area,
          meses
        };
      });

    return {
      inspecciones: inspecciones.actividades || [],
      capacitaciones: capacitaciones.actividades || [],
      reuniones: reunionesActividades,
      reunionesRaw: reunionesData.reuniones || [],
      entrenamientos: entrenamientos.actividades || [],
      anio
    };
  } catch (error) {
    Logger.log("Error en obtenerDatosPASSOCompleto: " + error.message);
    return { inspecciones: [], capacitaciones: [], reuniones: [], reunionesRaw: [], entrenamientos: [], error: error.message };
  }
}
