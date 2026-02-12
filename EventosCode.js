let cachedAccidentes = null;


function getSpreadsheetAccidentes() {
  if (!cachedAccidentes) {
    cachedAccidentes = SpreadsheetApp.openById("1Xo5HgaHfskg_mkguGTuuR_AcKeQUoRoG--ch-V8KTpw");
  }
  return cachedAccidentes;
}

function saveFormDataAccidentes(recordID, name, nombres, cargos, empresa, lugar, proceso, evento, tipo, area, comment, descanso1, descanso2, responsable, origen, detalle, estado, x, y) {
  const hojaevento = getSpreadsheetAccidentes().getSheetByName("B DATOS");
  const ultimaCol = 21;

  const ids = hojaevento.getRange(2, 1, hojaevento.getLastRow() - 1, 1).getValues().flat();
  const isEdit = recordID && recordID !== '';
  const index = isEdit ? ids.findIndex(id => String(id) === String(recordID)) : -1;

  let coordX = x;
  let coordY = y;

  if (isEdit && index !== -1) {
    const coordRango = hojaevento.getRange(index + 2, 20, 1, 2).getValues()[0];
    coordX = coordRango[0];
    coordY = coordRango[1];
  }

  const fecha1 = new Date(descanso1);
  const fecha2 = new Date(descanso2);
  const dias = Math.ceil((fecha2 - fecha1) / (1000 * 60 * 60 * 24)) + 1;

  const finalID = isEdit ? recordID : generateUniqueID();

  const fila = [
    finalID,
    new Date(), name, nombres, cargos, empresa, lugar, proceso, evento,
    tipo, area, comment, descanso1, descanso2, dias, responsable, origen,
    detalle, estado, coordX, coordY
  ];

  if (isEdit && index !== -1) {
    hojaevento.getRange(index + 2, 1, 1, ultimaCol).setValues([fila]);
  } else {
    hojaevento.appendRow(fila);
  }
}

function generateUniqueID() {
  const props = PropertiesService.getScriptProperties();
  const key = "ID_CONTADOR";

  let contador = parseInt(props.getProperty(key) || "100000", 10);
  contador++;
  props.setProperty(key, contador.toString());

  return contador.toString().padStart(6, '0');
}

// REEMPLAZO COMPLETO de searchByCoordinates
function searchByCoordinates(x, y) {
  const b = getSpreadsheetAccidentes().getSheetByName('B DATOS');
  const lastRow = b.getLastRow();
  const data = lastRow > 1 ? b.getRange(2, 1, lastRow - 1, 21).getValues() : [];

  const d = ensureDescansosSheet_();
  const dLast = d.getLastRow();
  const dVals = dLast >= 2 ? d.getRange(2, 1, dLast - 1, 11).getValues() : [];

  return data
    .filter(r => r[19] == x && r[20] == y)
    .map(r => {
      const id = r[0];
      const descansos = dVals
        .filter(dr => String(dr[0]) === String(id))
        .map(dr => ({
          inicio: Utilities.formatDate(new Date(dr[3]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          fin: Utilities.formatDate(new Date(dr[4]), Session.getScriptTimeZone(), 'yyyy-MM-dd')
        }));

      return {
        id: id,
        name: r[2],
        nombres: r[3],
        cargos: r[4],
        empresa: r[5],
        lugar: r[6],
        proceso: r[7],
        evento: r[8],
        tipo: r[9],
        area: r[10],
        comment: r[11],
        descanso1: Utilities.formatDate(r[12], Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        descanso2: Utilities.formatDate(r[13], Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        responsable: r[15],
        origen: r[16],
        detalle: r[17],
        estado: r[18],
        x: r[19],
        y: r[20],
        descansos: descansos
      };
    });
}


function getAllPoints(filter, tipoEvento, fechaInicioStr, fechaFinStr) {
  const sheet = getSpreadsheetAccidentes().getSheetByName("B DATOS");
  const lastRow = sheet.getLastRow();

  const hayFiltros = filter || tipoEvento || fechaInicioStr || fechaFinStr;
  const data = hayFiltros
    ? sheet.getRange(2, 1, lastRow - 1, 21).getValues().reverse()
    : (lastRow > 1
        ? sheet.getRange(Math.max(2, lastRow - 29), 1, Math.min(30, lastRow - 1), 21).getValues().reverse()
        : []);

  const fechaInicio = fechaInicioStr ? new Date(fechaInicioStr) : null;
  const fechaFin = fechaFinStr ? new Date(fechaFinStr) : null;

  const coincidencias = data.filter(r => {
    const nombreIncluye = !filter || r[2]?.toString().toLowerCase().includes(filter.toLowerCase()) || r[3]?.toString().toLowerCase().includes(filter.toLowerCase());
    const tipoCoincide = !tipoEvento || r[8]?.toString().toLowerCase().includes(tipoEvento.toLowerCase());

    let fechaEvento = r[12];
    if (!(fechaEvento instanceof Date)) {
      try {
        fechaEvento = new Date(fechaEvento);
      } catch (e) {
        fechaEvento = null;
      }
    }

    const fechaValida = (!fechaInicio || (fechaEvento && fechaEvento >= fechaInicio)) &&
                        (!fechaFin || (fechaEvento && fechaEvento <= fechaFin));

    return nombreIncluye && tipoCoincide && fechaValida;
  });

  return coincidencias.map(r => [
    r[19], r[20], r[0], r[3], r[8],
    `${Utilities.formatDate(r[12], Session.getScriptTimeZone(), 'yyyy-MM-dd')} al ${Utilities.formatDate(r[13], Session.getScriptTimeZone(), 'yyyy-MM-dd')} (${r[14]} días)`,
    r[18]
  ]);
}

function deleteByIDAccindentes(id) {
  const hojaevento = getSpreadsheetAccidentes().getSheetByName("B DATOS");
  const datos = hojaevento.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(id)) {
      hojaevento.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

let empleadosCache = null;

function getNombreEnferno(idEnfermo) {
  if (!empleadosCache) {
    const hojaevento = getSpreadsheetPersonal().getSheetByName('PERSONAL');
    const lastRow = hojaevento.getLastRow();
    if (lastRow < 2) return { nombre: "", cargo: "", empresa: "" };

    const data = hojaevento.getRange(2, 2, lastRow - 1, 6).getValues();
    empleadosCache = new Map(
      data.map(r => [
        r[0],
        {
          nombre: r[1] || "",
          empresa: r[3] || "",
          cargo: r[5] || ""
        }
      ])
    );
  }

  return empleadosCache.get(idEnfermo) || { nombre: "", cargo: "", empresa: "" };
}

function getParte() {
  const sheet = getSpreadsheetAccidentes().getSheetByName("Listas");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { parte: [], atencion: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // Columnas A y B

  const atencion = data.map(r => r[0]).filter(String); // Columna A
  const parte = data.map(r => r[1]).filter(String);    // Columna B

  return { parte, atencion };
}


function getTiposEvento() {
  const hojaevento = getSpreadsheetAccidentes().getSheetByName("B DATOS");
  const ultimaFila = hojaevento.getLastRow();
  const datos = hojaevento.getRange(2, 9, ultimaFila - 1, 1).getValues().flat();

  const tipos = [...new Set(datos.filter(Boolean))];
  return tipos.sort();
}


let stockCache = null;

function getStockData() {
  if (!stockCache) {
    const hojaevento = getSpreadsheetAccidentes().getSheetByName('Stock');
    const lastRow = hojaevento.getLastRow();
    stockCache = lastRow < 2 ? [] : hojaevento.getRange(2, 1, lastRow - 1, 6).getValues();
  }
  return stockCache;
}

function puxaProdutos(produto) {
  const data = getStockData(); // ✅ Usa caché
  return data
    .filter(r => r[0] === produto)
    .map(r => ({
      produto: r[0],
      nome: r[1] || "",
      custo: parseFloat(r[2]) ? parseFloat(r[2]).toFixed(2) : "0.00",
      valor: parseFloat(r[3]) ? parseFloat(r[3]).toFixed(2) : "0.00"
    }));
}


function salvar(pedido) {
  if (!pedido || !pedido.length) return;

  const hojaevento = getSpreadsheetAccidentes().getSheetByName("Salidas Medicamentos");
  const lastRow = hojaevento.getLastRow();
  const lastNumero = parseInt(hojaevento.getRange(lastRow, 1).getValue(), 10);
  const nuevoNumero = isNaN(lastNumero) ? 1 : lastNumero + 1;

  const timestamp = new Date();
  const rows = pedido.map(item => [
    nuevoNumero,
    item.cliente || "",
    item.nome || "",
    item.quantidade || 0,
    item.custo || 0,
    item.valor || 0,
    timestamp
  ]);

  hojaevento.getRange(lastRow + 1, 1, rows.length, 7).setValues(rows);
}

function getMedicamentos() {
  const data = getStockData();
  return [...new Set(data.map(r => r[1]).filter(String))];
}

function getMedicamentoData(medicamento) {
  const data = getStockData();
  const found = data.find(r => r[1] === medicamento);
  if (!found) return {};

  return {
    grupo: found[0],
    costo: found[2],
    proveedor: found[4],
    presentacion: found[5]
  };
}

function resetStockCache() {
  stockCache = null;
}

// NUEVO
function ensureDescansosSheet_() {
  const ss = getSpreadsheetAccidentes();
  const sh = ss.getSheetByName('DESCANSOS') || ss.insertSheet('DESCANSOS');
  if (sh.getLastRow() === 0) {
    // Incluyo fechas porque son indispensables por cada descanso
    sh.getRange(1, 1, 1, 11).setValues([[
      'ID','DNI','Descripción','Fecha inicial','Fecha fin','Dias de descanso',
      'Responsable','Atención','Link o detalles','Estado','Timestamp'
    ]]);
  }
  return sh;
}
// NUEVO
function saveFormDataAccidentesV2(req) {
  const ss = getSpreadsheetAccidentes();
  const b = ss.getSheetByName('B DATOS');
  const d = ensureDescansosSheet_();
  const TZ = Session.getScriptTimeZone();

  // ===== Helpers =====
  const norm = (val) => {
    if (val instanceof Date) return Utilities.formatDate(val, TZ, 'yyyy-MM-dd');
    const s = String(val || '');
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    const dt = new Date(s);
    return isNaN(dt) ? '' : Utilities.formatDate(dt, TZ, 'yyyy-MM-dd');
  };
  const toDateLocal = (ymd) => new Date(`${ymd}T00:00:00`);
  const daysInclusive = (iniY, finY) => {
    const a = toDateLocal(iniY), bdt = toDateLocal(finY);
    return Math.max(Math.floor((bdt - a) / 86400000) + 1, 0);
  };
  const overlaps = (a, b) => toDateLocal(a.inicio) <= toDateLocal(b.fin) && toDateLocal(b.inicio) <= toDateLocal(a.fin);
  const byStart = (a, b) => a.inicio < b.inicio ? -1 : a.inicio > b.inicio ? 1 : 0;

  // ===== Identidad / edición =====
  const isEdit = !!(req.recordID && String(req.recordID).trim() !== '');
  const ids = b.getRange(2, 1, Math.max(b.getLastRow() - 1, 0), 1).getValues().flat();
  const idx = isEdit ? ids.findIndex(id => String(id) === String(req.recordID)) : -1;
  const finalID = isEdit ? req.recordID : generateUniqueID();

  // Mantener X,Y si edición
  let coordX = req.x, coordY = req.y;
  if (isEdit && idx !== -1) {
    const xy = b.getRange(idx + 2, 20, 1, 2).getValues()[0];
    coordX = xy[0]; coordY = xy[1];
  }

  // ===== Incoming: normalizar y deduplicar =====
  const incomingRaw = (req.descansos || [])
    .map(r => {
      let ini = norm(r.inicio), fin = norm(r.fin);
      if (ini && fin && fin < ini) { const t = ini; ini = fin; fin = t; }
      return { inicio: ini, fin: fin };
    })
    .filter(r => r.inicio && r.fin);

  const seen = new Set();
  const incoming = [];
  for (const r of incomingRaw) {
    const k = `${r.inicio}|${r.fin}`;
    if (!seen.has(k)) { seen.add(k); incoming.push(r); }
  }

  // Validación: solape interno en lo que viene del front
  const sortedIncoming = incoming.slice().sort(byStart);
  for (let i = 1; i < sortedIncoming.length; i++) {
    if (overlaps(sortedIncoming[i-1], sortedIncoming[i])) {
      return { error: 'OVERLAP_SELF', message: 'Hay rangos de descanso que se superponen entre sí.', rangos: sortedIncoming };
    }
  }

  // ===== Leer EXISTENTES de DESCANSOS (con fila real) =====
  const dLast = d.getLastRow();
  const dVals = dLast >= 2 ? d.getRange(2, 1, dLast - 1, 11).getValues() : [];
  const rows = dVals.map((r, i) => ({
    sheetRow: i + 2,   // fila real en la hoja
    id: r[0], dni: r[1], desc: r[2],
    ini: r[3], fin: r[4], dias: r[5],
    resp: r[6], aten: r[7], link: r[8], estado: r[9],
    ts: r[10]
  }));

  const existingForID = rows
    .filter(o => String(o.id) === String(finalID))
    .map(o => ({
      sheetRow: o.sheetRow,
      inicio: norm(o.ini),
      fin: norm(o.fin),
      dias: Number(o.dias) || daysInclusive(norm(o.ini), norm(o.fin)),
      ts: (o.ts instanceof Date) ? o.ts : (o.ts ? new Date(o.ts) : null)
    }));

  // Última fila (por timestamp; si empate, por número de fila)
  let lastRowForID = null;
  if (existingForID.length) {
    lastRowForID = existingForID
      .slice()
      .sort((a,b) => {
        const at = a.ts ? a.ts.getTime() : 0;
        const bt = b.ts ? b.ts.getTime() : 0;
        if (at === bt) return a.sheetRow - b.sheetRow;
        return at - bt;
      })[existingForID.length - 1];
  }

  // ===== Diferencias: qué agregar y qué borrar =====
  const incomingKeys = new Set(incoming.map(r => `${r.inicio}|${r.fin}`));
  const existingKeys = new Set(existingForID.map(r => `${r.inicio}|${r.fin}`));

  const toAppend = incoming.filter(r => !existingKeys.has(`${r.inicio}|${r.fin}`));
  const toDelete = existingForID.filter(r => !incomingKeys.has(`${r.inicio}|${r.fin}`));

  // Validación: que lo nuevo no se solape con lo que se quedará (existing - toDelete)
  const remainingExisting = existingForID.filter(r => incomingKeys.has(`${r.inicio}|${r.fin}`)); // los que permanecen
  for (const inc of toAppend) {
    for (const ex of remainingExisting) {
      if (overlaps(inc, ex)) {
        return {
          error: 'OVERLAP_EXISTING',
          message: 'El nuevo rango se superpone con un descanso ya registrado.',
          conflicto: { nuevo: inc, existente: { inicio: ex.inicio, fin: ex.fin } }
        };
      }
    }
  }

  // ===== Aplicar cambios en DESCANSOS =====
  // 1) Borrar rangos que el usuario quitó
  if (toDelete.length > 0) {
    const rowsToDelete = toDelete
      .map(r => r.sheetRow)
      .sort((a,b) => b - a); // borrar de abajo hacia arriba
    rowsToDelete.forEach(rn => d.deleteRow(rn));
  }

  // 2) Agregar rangos nuevos
  let appended = 0;
  if (toAppend.length > 0) {
    const rowsIns = toAppend.map(r => [
      finalID,
      req.name || '',                       // DNI
      req.comment || req.descripcion || '', // Descripción
      toDateLocal(r.inicio),                // Fecha inicial
      toDateLocal(r.fin),                   // Fecha fin
      daysInclusive(r.inicio, r.fin),       // Días
      req.responsable || '',                // Responsable
      req.atencion || req.origen || '',     // Atención
      req.detalle || '',                    // Link o detalles
      req.estado || '',                     // Estado
      new Date()                            // Timestamp
    ]);
    d.getRange(d.getLastRow() + 1, 1, rowsIns.length, 11).setValues(rowsIns);
    appended = rowsIns.length;
  }

  // 3) Si NO hubo cambios de rangos (ni agrega ni borra): actualizar SOLO metadatos de la última fila
  if (toAppend.length === 0 && toDelete.length === 0 && lastRowForID) {
    const r0 = lastRowForID.sheetRow;
    d.getRange(r0, 2, 1, 1).setValue(req.name || '');            // DNI
    d.getRange(r0, 3, 1, 1).setValue(req.comment || req.descripcion || ''); // Descripción
    d.getRange(r0, 7, 1, 1).setValue(req.responsable || '');     // Responsable
    d.getRange(r0, 8, 1, 1).setValue(req.atencion || req.origen || ''); // Atención
    d.getRange(r0, 9, 1, 1).setValue(req.detalle || '');         // Link/detalles
    d.getRange(r0, 10, 1, 1).setValue(req.estado || '');         // Estado
    d.getRange(r0, 11, 1, 1).setValue(new Date());               // Timestamp
  }

  // ===== Recalcular agregados para B DATOS =====
  // Releer todo para este ID (ya con borrados/insertados aplicados)
  const dLast2 = d.getLastRow();
  const vals2 = dLast2 >= 2 ? d.getRange(2, 1, dLast2 - 1, 11).getValues() : [];
  const allForID = vals2
    .filter(r => String(r[0]) === String(finalID))
    .map(r => ({
      inicio: norm(r[3]),
      fin: norm(r[4]),
      dias: Number(r[5]) || daysInclusive(norm(r[3]), norm(r[4]))
    }));

  const totalDias = allForID.reduce((a, r) => a + (r.dias || 0), 0);
  const minIni = allForID.length ? new Date(Math.min.apply(null, allForID.map(r => toDateLocal(r.inicio)))) : '';
  const maxFin = allForID.length ? new Date(Math.max.apply(null, allForID.map(r => toDateLocal(r.fin)))) : '';

  // Actualizar/insertar fila en B DATOS
  const fila = [
    finalID,
    new Date(),                 // Date (ahora)
    req.name,                   // DNI
    req.nombres,                // Nombres y Apellido
    req.cargos,                 // Cargo
    req.empresa,                // Empresa
    req.lugar,                  // Lugar
    req.proceso,                // Proceso
    req.evento,                 // Evento
    req.tipo,                   // Tipo
    req.area,                   // Parte / Área
    req.comment,                // Descripción
    minIni || '',               // Fecha inicial (mín)
    maxFin || '',               // Fecha fin (máx)
    totalDias,                  // Días de descanso (suma)
    req.responsable,            // Responsable
    req.origen,                 // Atención
    req.detalle,                // Link o detalles
    req.estado,                 // Estado
    coordX, coordY              // X, Y
  ];

  if (isEdit && idx !== -1) b.getRange(idx + 2, 1, 1, 21).setValues([fila]);
  else b.appendRow(fila);

  return {
    ok: true,
    appended,
    deleted: toDelete.length,
    onlyMetaUpdated: (toAppend.length === 0 && toDelete.length === 0 && !!lastRowForID),
    totalDias,
    fechaInicial: minIni,
    fechaFin: maxFin
  };
}




