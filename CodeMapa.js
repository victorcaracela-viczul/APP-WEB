
let cachedMapaRiesgos = null;
function getSpreadsheetMapaRiesgos() {
  if (!cachedMapaRiesgos) {
    cachedMapaRiesgos = SpreadsheetApp.openById("1EfQvY59m1l1SB_GD__CzL-qJQFdbtYzM9Y2q1u2L3cI"); // HOJA DE CALCULO MAPA DE RIESGOS
  }
  return cachedMapaRiesgos;
}
 const carpetaIdMapa = '1dwtsWDNfgsYKkSc7wCAILROVQg5uTMwJ'; // CARPETA ICONOS
/***********************
 * CONFIG
 ***********************/
const hojaMAPAS  = "MAPAS";
const hojaICONOS = "ICONOS";

/***********************
 * UTILS
 ***********************/
function ensureCols_() {
  let hoja = getSpreadsheetMapaRiesgos().getSheetByName(hojaMAPAS);
  if (!hoja) return;

  // A:ID B:Titulo C:Fondo D:Textos E:Iconos F:FondoWidth G:FondoHeight H:FondoScale I:Fecha J:Autor K:Password
  const neededHeaders = [
    "ID","Titulo","Fondo","Textos","Iconos",
    "FondoWidth","FondoHeight","FondoScale","Fecha","Autor","Password"
  ];

  const lastCol = Math.max(hoja.getLastColumn(), neededHeaders.length);
  let header = hoja.getRange(1, 1, 1, lastCol).getValues()[0] || [];

  // Extiende si hace falta
  if (header.length < neededHeaders.length) {
    hoja.insertColumnsAfter(header.length || 1, neededHeaders.length - header.length);
    header = hoja.getRange(1, 1, 1, neededHeaders.length).getValues()[0] || [];
  }

  // Escribe encabezados faltantes o incorrectos
  neededHeaders.forEach((name, i) => {
    if (!header[i] || header[i] !== name) hoja.getRange(1, i + 1).setValue(name);
  });
}


function mapRowToObj_(row){
  return {
    id:          row[0]?.toString() || "",
    titulo:      row[1] || "",
    fondo:       row[2] || "",
    textos:      JSON.parse(row[3] || "[]"),
    iconos:      JSON.parse(row[4] || "[]"),
    fondoWidth:  row[5] || null,
    fondoHeight: row[6] || null,
    fondoScale:  row[7] || null,
    fecha:       row[8] ? new Date(row[8]).toISOString() : null,
    autor:       row[9] || "",
    pass:        row[10] || ""   // <-- K: Password
  };
}


function norm_(s){
  return (s||"")
    .toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g,"");
}

// Extrae ID de un URL tipo https://lh5.googleusercontent.com/d/<ID>  o drive/uc?id=<ID>
function extractDriveId_(urlOrId){
  if (!urlOrId) return "";
  if (/^[A-Za-z0-9_\-]{20,}$/.test(urlOrId)) return urlOrId; // ya es un ID
  const m1 = /\/d\/([A-Za-z0-9_\-]+)/.exec(urlOrId);
  if (m1) return m1[1];
  const m2 = /[?&]id=([A-Za-z0-9_\-]+)/.exec(urlOrId);
  if (m2) return m2[1];
  return "";
}

/***********************
 * MAPAS: GET (completo)
 ***********************/
function getMapas() {
  let hoja = getSpreadsheetMapaRiesgos().getSheetByName(hojaMAPAS);

  if (!hoja) {
    hoja = getSpreadsheetMapaRiesgos().insertSheet(hojaMAPAS);
    hoja.getRange(1, 1, 1, 11).setValues([[
      "ID","Titulo","Fondo","Textos","Iconos",
      "FondoWidth","FondoHeight","FondoScale","Fecha","Autor","Password"
    ]]);
    return [];
  }
  ensureCols_();

  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return [];

  const maps = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    maps.push(mapRowToObj_(row));
  }
  return maps;
}


/*************************************
 * MAPAS: GET paginado + búsqueda
 *  - search: string
 *  - offset: number (0,20,40…)
 *  - limit : number (ej. 20)
 *************************************/
function getMapasPage(search, offset, limit){
  let hoja   = getSpreadsheetMapaRiesgos().getSheetByName(hojaMAPAS);
  if(!hoja){
    hoja = getSpreadsheetMapaRiesgos().insertSheet(hojaMAPAS);
    hoja.getRange(1,1,1,11).setValues([[
      "ID","Titulo","Fondo","Textos","Iconos",
      "FondoWidth","FondoHeight","FondoScale","Fecha","Autor","Password"
    ]]);
    return { items: [], total: 0 };
  }
  ensureCols_();

  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return { items: [], total: 0 };

  const term = norm_(search || "");
  const filtered = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    if (!row[0]) continue;
    if (!term){
      filtered.push(row);
      continue;
    }
    const tituloOk = norm_(row[1]).includes(term);
    const autorOk  = norm_(row[9]).includes(term);
    let textosOk = false;
    if (!(tituloOk || autorOk)){
      const textosStr = (row[3] || "").toString().toLowerCase();
      textosOk = textosStr.includes(term);
    }
    if (tituloOk || autorOk || textosOk) filtered.push(row);
  }

  filtered.sort((a,b)=>{
    const da = a[8] ? new Date(a[8]).getTime() : 0;
    const db = b[8] ? new Date(b[8]).getTime() : 0;
    return db - da;
  });

  const total = filtered.length;
  const start = Math.max(0, +offset|0);
  const end   = Math.min(start + (+limit || 20), total);
  const page  = filtered.slice(start, end).map(mapRowToObj_);

  return { items: page, total };
}


/***********************
 * MAPAS: SAVE
 ***********************/
function saveMapa(mapa) {
  let hoja = getSpreadsheetMapaRiesgos().getSheetByName(hojaMAPAS);

  if (!hoja) {
    hoja = getSpreadsheetMapaRiesgos().insertSheet(hojaMAPAS);
    hoja.getRange(1, 1, 1, 11).setValues([[
      "ID","Titulo","Fondo","Textos","Iconos",
      "FondoWidth","FondoHeight","FondoScale","Fecha","Autor","Password"
    ]]);
  }
  ensureCols_();

  const data = hoja.getDataRange().getValues();
  const ids  = data.map(r => r[0]?.toString());
  let idx    = ids.indexOf(mapa.id); // índice base 0 sobre "data"
  let fila; // 1-based para setValues

  let prevPass = "";
  if (idx === -1) {
    fila = hoja.getLastRow() + 1;
  } else {
    fila = idx + 1;
    prevPass = data[idx][10] || ""; // K: Password
  }

  // Fondo a URL completa si solo es ID
  let fondoURL = mapa.fondo;
  if (fondoURL && !/^https?:\/\//i.test(fondoURL)) {
    fondoURL = `https://lh5.googleusercontent.com/d/${fondoURL}`;
  }

  const passToSave = (typeof mapa.pass === 'string' && mapa.pass.trim().length)
    ? mapa.pass.trim()
    : prevPass;

  const now = new Date();
  hoja.getRange(fila, 1, 1, 11).setValues([[
    mapa.id,
    mapa.titulo || "",
    fondoURL || "",
    JSON.stringify(mapa.textos || []),
    JSON.stringify(mapa.iconos || []),
    mapa.fondoWidth || "",
    mapa.fondoHeight || "",
    mapa.fondoScale || "",
    now,
    mapa.autor || "",
    passToSave || ""
  ]]);

  return true;
}


/***********************
 * MAPAS: DELETE
 ***********************/
function deleteMapa(id) {
  const hoja = getSpreadsheetMapaRiesgos().getSheetByName(hojaMAPAS);
  if (!hoja) return false;

  const data = hoja.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === id) {
      hoja.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

/***********************
 * ÍCONOS: GET
 ***********************/
function getIconos() {
  let hoja = getSpreadsheetMapaRiesgos().getSheetByName(hojaICONOS);
  if (!hoja) {
    hoja = getSpreadsheetMapaRiesgos().insertSheet(hojaICONOS);
  }
  ensureIconCols_();

  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return [];

  const icons = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    icons.push({
      id:    row[0].toString(),
      tipo:  row[1] || '',
      nombre:row[2] || '',
      img:   row[3] || ''
    });
  }
  return icons;
}


/***********************
 * ÍCONOS: SUBIR
 ***********************/
function subirIcono(tipo, nombre, base64) {
  try {
    const imageId  = subirImagenADrive((nombre || 'icon') + '_icon', base64);
    const imageUrl = `https://lh5.googleusercontent.com/d/${imageId}`;

    let hoja = getSpreadsheetMapaRiesgos().getSheetByName(hojaICONOS);
    if (!hoja) hoja = getSpreadsheetMapaRiesgos().insertSheet(hojaICONOS);
    ensureIconCols_(); // asegura: ID | Tipo | Nombre | URL

    const id = Utilities.getUuid();
    const nuevaFila = hoja.getLastRow() + 1;
    hoja.getRange(nuevaFila, 1, 1, 4).setValues([[id, tipo || '', nombre || '', imageUrl]]);
    return true;
  } catch (error) {
    console.error('Error subiendo ícono:', error);
    throw new Error('No se pudo subir el ícono: ' + error.message);
  }
}


/***********************
 * ÍCONOS: ELIMINAR por ID
 ***********************/
function eliminarIcono(iconoId) {
  try {
    const hoja = getSpreadsheetMapaRiesgos().getSheetByName(hojaICONOS);
    if (!hoja) return false;

    ensureIconCols_();
    const data = hoja.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === iconoId) {
        try {
          const imageUrlOrId = data[i][3]; // D: URL
          const fileId = extractDriveId_(imageUrlOrId);
          if (fileId) DriveApp.getFileById(fileId).setTrashed(true);
        } catch (driveError) {
          console.log('No se pudo eliminar la imagen de Drive:', driveError);
        }
        hoja.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  } catch (error) {
    console.error('Error eliminando ícono:', error);
    throw new Error('No se pudo eliminar el ícono: ' + error.message);
  }
}

/**********************************
 * ÍCONOS: ELIMINAR por NOMBRE (legacy)
 **********************************/
function eliminarIconoPorNombre(nombre) {
  try {
    const hoja = getSpreadsheetMapaRiesgos().getSheetByName(hojaICONOS);
    if (!hoja) return false;

    ensureIconCols_();
    const data = hoja.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString() === nombre) { // C: Nombre
        try {
          const imageUrlOrId = data[i][3]; // D: URL
          const fileId = extractDriveId_(imageUrlOrId);
          if (fileId) DriveApp.getFileById(fileId).setTrashed(true);
        } catch (driveError) {
          console.log('No se pudo eliminar la imagen de Drive:', driveError);
        }
        hoja.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  } catch (error) {
    console.error('Error eliminando ícono por nombre:', error);
    throw new Error('No se pudo eliminar el ícono: ' + error.message);
  }
}


/***********************
 * FONDOS: SUBIR
 ***********************/
function subirFondo(nombre, base64) {
  return subirImagenADrive(nombre, base64);
}

/***********************
 * DRIVE: subir imagen
 ***********************/
function subirImagenADrive(nombre, base64) {
  try {
   
    const carpeta   = DriveApp.getFolderById(carpetaIdMapa);

    const base64Clean = base64.includes(',') ? base64.split(',')[1] : base64;
    const contenido   = Utilities.base64Decode(base64Clean);

    let mimeType = MimeType.JPEG;
    if (base64.includes('data:image/png')) mimeType = MimeType.PNG;
    else if (base64.includes('data:image/gif')) mimeType = MimeType.GIF;

    const blob    = Utilities.newBlob(contenido, mimeType, nombre);
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getId();
  } catch (error) {
    console.error('Error subiendo imagen a Drive:', error);
    throw new Error('No se pudo subir la imagen: ' + error.message);
  }
}

/***********************
 * TEST
 ***********************/
function testStructure() {
  try {
    ensureCols_();
    ensureIconCols_();
    const m = getMapas();
    const i = getIconos();
    return { mapas: m.length, iconos: i.length };
  } catch (error) {
    console.error('Error en test:', error);
    throw error;
  }
}

function ensureIconCols_() {
  let hoja = getSpreadsheetMapaRiesgos().getSheetByName(hojaICONOS);
  if (!hoja) {
    hoja = getSpreadsheetMapaRiesgos().insertSheet(hojaICONOS);
  }
  // Queremos: ID | Tipo | Nombre | URL
  const needed = ["ID","Tipo","Nombre","URL"];
  const lastCol = Math.max(hoja.getLastColumn(), needed.length);
  let header = hoja.getRange(1, 1, 1, lastCol).getValues()[0] || [];

  // Si la hoja estaba con 3 columnas antiguas (ID, Nombre, Imagen), la “migramos” rápido:
  if (header.length === 3 && String(header[1]).toLowerCase().includes('nombre')) {
    hoja.insertColumnAfter(1); // inserta "Tipo" en col 2
    header = hoja.getRange(1, 1, 1, 4).getValues()[0] || [];
  }

  // Extiende ancho si falta
  if (header.length < needed.length) {
    hoja.insertColumnsAfter(header.length || 1, needed.length - header.length);
    header = hoja.getRange(1, 1, 1, needed.length).getValues()[0] || [];
  }

  // Escribe encabezados
  needed.forEach((name, i) => {
    if (!header[i] || header[i] !== name) hoja.getRange(1, i + 1).setValue(name);
  });
}