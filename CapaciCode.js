//Inicio Capacitaciones                                                                                                                   
    //let ss = SpreadsheetApp.openById("1MyXsN09Jf23dcimniDLrDu3luDgasCk8EbLHOp3gzGw")                                                        
    const foldercharlas = "1IWmNW4wMZbC43QHrdivlRtAwnTHa3c2v"; //CARPETA DE CHARLAS                                            
     const foldefirmascap = "1GcIoeFFtpZ6EISt0w5R1byNkDy5I3ugi";  //CARPETA DE FIRMA CAPCITACIONES                                                                                          
    let cachedCapacitaciones = null;                                                                                                          
    function getSpreadsheetCapacitaciones() {                                                                                                 
      if (!cachedCapacitaciones) {                                                                                                            
        cachedCapacitaciones = SpreadsheetApp.openById("1Ev5_B3jMtjy_xXt13NYBXYwFA-maFAeLSKfiCFIsMQo"); //HOJA DE CALCULO CAPACITACIONES                                       
      }                                                                                                                                       
      return cachedCapacitaciones;                                                                                                            
    }                                                                                                                                         
                                                                                                                                              
function cargarDatosGlobales() {
  const ssCap = getSpreadsheetCapacitaciones();
  const ssPersonal = getSpreadsheetPersonal();
  const hojaMatriz = ssCap.getSheetByName("Matriz");
  const hojaPersonal = ssPersonal.getSheetByName("PERSONAL");
  const hojaBDatos = ssCap.getSheetByName("B DATOS");
  const lastRowMatriz = hojaMatriz.getLastRow();
  const lastColMatriz = hojaMatriz.getLastColumn();
  const matrizDatos = hojaMatriz.getRange(1, 1, lastRowMatriz, lastColMatriz).getValues();

  // 🔹 Filas horizontales actualizadas (2 filas nuevas: Tipo de Programa y Responsable)
  const cursos = matrizDatos[13].slice(4); // Fila 14 → Temas
  const cargos = matrizDatos.slice(14).map(f => f[3]); // Columna D desde fila 15 en adelante
  const matriz = matrizDatos.slice(14).map(f => f.slice(4, 4 + cursos.length));
  const lastRowPersonal = hojaPersonal.getLastRow();
  const personalDatos = hojaPersonal.getRange(1, 1, lastRowPersonal, 7).getValues();
  const personalCargos = personalDatos.slice(1).map(f => f[6]);
  const lastRowBDatos = hojaBDatos.getLastRow();
  const bdDatos = hojaBDatos.getRange(1, 1, lastRowBDatos, 13).getValues();

  return {
    matrizDatos,
    cursos,
    cargos,
    matriz,
    personalDatos,
    personalCargos,
    bdDatos

  };

}
                                                                                                                                     
 function obtenerDatosPorDNI(dni) {
  const datos = cargarDatosGlobales();
  const { personalDatos, matrizDatos, bdDatos } = datos;

  // 🔹 Buscar persona por DNI
  let persona = null;
  for (let i = 1; i < personalDatos.length; i++) {
    if (personalDatos[i][1].toString().trim() === dni.toString().trim()) {
      persona = {
        nombre: personalDatos[i][2],
        empresa: personalDatos[i][4],
        cargo: personalDatos[i][6]
      };
      break;
    }
  }

  if (!persona) return { encontrado: false };

  // 🔹 Extraer filas horizontales de la matriz (+2 por Tipo de Programa y Responsable)
  const tiposProg = matrizDatos[1].slice(4);       // Fila 2 → Tipo de Programa (NUEVA)
  const responsables = matrizDatos[2].slice(4);     // Fila 3 → Responsable (NUEVA)
  const links = matrizDatos[3].slice(4);            // Fila 4 → Link
  const programaciones = matrizDatos[4].slice(4);   // Fila 5 → Programación
  const temporalidades = matrizDatos[5].slice(4);   // Fila 6 → Vigencia
  const imagen = matrizDatos[6].slice(4);           // Fila 7 → Imagen
  const puntajesMin = matrizDatos[7].slice(4).map(p => parseFloat(p)); // Fila 8 → Puntaje
  const duraciones = matrizDatos[8].slice(4).map(d => parseInt(d));    // Fila 9 → Duración
  const capacitadores = matrizDatos[9].slice(4);    // Fila 10 → Capacitador
  const areas = matrizDatos[10].slice(4);           // Fila 11 → Área
  const horasLectivas = matrizDatos[11].slice(4);   // Fila 12 → Horas lectivas
  const tieneCertificacion = matrizDatos[12].slice(4); // Fila 13 → Certificación
  const cursos = matrizDatos[13].slice(4);          // Fila 14 → Temas

  // 🔹 Buscar la fila del cargo
  const filaCursos = matrizDatos.find(f => f[3]?.toString().toLowerCase().trim() === persona.cargo.toLowerCase().trim());
  if (!filaCursos) return { encontrado: true, persona, cursos: [] };

  // 🔹 Indexar evaluaciones
  const evaluacionesMap = {};
  for (let i = 1; i < bdDatos.length; i++) {
    const [dniBD, , , , tema, , puntaje, fecha] = bdDatos[i];
    if (!evaluacionesMap[dniBD]) evaluacionesMap[dniBD] = {};
    if (!evaluacionesMap[dniBD][tema]) evaluacionesMap[dniBD][tema] = [];
    evaluacionesMap[dniBD][tema].push({ puntaje: parseFloat(puntaje), fecha: new Date(fecha) });
  }

  const hoy = new Date();
  const cursosEncontrados = [];

  for (let j = 4; j < filaCursos.length; j++) {
    if (filaCursos[j] !== true && filaCursos[j] !== "VERDADERO") continue;

    const temaCurso = cursos[j - 4];
    const evaluaciones = evaluacionesMap[dni]?.[temaCurso] || [];

    const intentosHoy = evaluaciones.filter(ev => new Date(ev.fecha).toDateString() === hoy.toDateString()).length;

    // Mejor puntaje y última fecha
    let mejorPuntaje = null;
    let fechaEvaluacion = null;
    for (const ev of evaluaciones) {
      if (
        mejorPuntaje === null ||
        ev.puntaje > mejorPuntaje ||
        (ev.puntaje === mejorPuntaje && ev.fecha > fechaEvaluacion)
      ) {
        mejorPuntaje = ev.puntaje;
        fechaEvaluacion = ev.fecha;
      }
    }

    // Estado
    let estadoCurso = "Pendiente";
    if (mejorPuntaje !== null && fechaEvaluacion instanceof Date && !isNaN(fechaEvaluacion)) {
      const vencimiento = new Date(fechaEvaluacion);
      vencimiento.setDate(vencimiento.getDate() + temporalidades[j - 4]);
      estadoCurso = hoy > vencimiento
        ? "Caducado"
        : mejorPuntaje >= puntajesMin[j - 4] ? "Aprobado" : "Reprobado";
    }

    // Programación
    const prog = programaciones[j - 4];
    let textoProgramacion = "";
    let fechaProg = null;

    if (prog instanceof Date && !isNaN(prog)) {
      fechaProg = prog;
    } else if (typeof prog === "string" && prog.trim() !== "") {
      const match = prog.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s*(\d{1,2}:\d{2}(?::\d{2})?)?/);
      if (match) {
        const [_, d, m, y, hms] = match;
        fechaProg = new Date(`${m}/${d}/${y} ${hms || "00:00"}`);
      }
    }

    if (fechaProg instanceof Date && !isNaN(fechaProg)) {
      const fin = new Date(fechaProg.getTime() + duraciones[j - 4] * 60000);
      if (hoy >= fechaProg && hoy <= fin) {
        const fechaOculta = Utilities.formatDate(fechaProg, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        textoProgramacion = `En vivo ((🔴))|${fechaOculta}`;
      } else if (hoy < fechaProg) {
        textoProgramacion = Utilities.formatDate(fechaProg, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      } else {
        textoProgramacion = "No previsto";
      }
    } else {
      textoProgramacion = prog ? prog.toString() : "";
    }

    // 🔹 Agregar al resultado
    cursosEncontrados.push({
      tema: temaCurso,
      tipoProg: tiposProg[j - 4] || '',
      responsable: responsables[j - 4] || '',
      link: links[j - 4],
      programacion: textoProgramacion,
      image: imagen[j - 4],
      puntajeMinimo: puntajesMin[j - 4],
      duracion: duraciones[j - 4],
      horasLectivas: horasLectivas[j - 4],
      temporabilidad: temporalidades[j - 4],
      tieneCertificacion: (tieneCertificacion[j - 4] === true || String(tieneCertificacion[j - 4]).toUpperCase() === "VERDADERO"),
      puntaje: mejorPuntaje !== null ? mejorPuntaje : "-",
      area: areas[j - 4],
      capacitador: capacitadores[j - 4],
      estado: estadoCurso,
      intentosHoy,
      fecha: fechaEvaluacion
        ? Utilities.formatDate(fechaEvaluacion, Session.getScriptTimeZone(), "dd/MM/yyyy")
        : "-"
    });
  }

  return { encontrado: true, persona, cursos: cursosEncontrados };
}
                                                                                                                                         
                                                                                                                                              
    //QUIZZ                                                                                                                                   
    /** Nombre hojas */                                                                                                                       
    var quizData = "Examen"; //                                                                                                               
    var bd = "B DATOS"; //                                                                                                                    
                                                                                                                                              
    /** ******** Cuestionarios  ******** **/                                                                                                  
    function getDataQuestion(selectedValue) {                                                                                                 
      const sheet = getSpreadsheetCapacitaciones().getSheetByName(quizData);                                                                  
      const lastRow = sheet.getLastRow();                                                                                                     
      const data = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues() // A2:K                                                           
        .filter(d => d[0] !== "" && d[2] === selectedValue); // d[2] = Tema                                                                   
                                                                                                                                              
      const maxQ = data.length;                                                                                                               
      const correctAnswer = [];                                                                                                               
      const pointValues = [];                                                                                                                 
                                                                                                                                              
      const radioLists = data.map((d, index) => {                                                                                             
        const pregunta = d[3];                                                                                                                
        const urlImagen = d[4];                                                                                                               
        const opciones = [d[5], d[6], d[7], d[8]];                                                                                            
        const correcta = d[9]; // valor 1-4                                                                                                   
        const puntos = parseFloat(d[10]) || 0;                                                                                                
        const id = index + 1;                                                                                                                 
                                                                                                                                              
        correctAnswer.push(correcta);                                                                                                         
        pointValues.push(puntos);                                                                                                             
                                                                                                                                              
        let imgHtml = urlImagen ? `<img class="img-fluid cat mt-2 mb-3" src="${urlImagen}" alt="imagen">` : "";                               
                                                                                                                                              
        return `                                                                                                                              
          <div id="${id}" class="fade-in-page mt-4" style="display:none">                                                                     
            <hr>                                                                                                                              
            <div class="row mt-2">                                                                                                            
              <label class="radio-label mt-2">                                                                                                
                <div><span class="inner-label1">${pregunta}</span></div>                                                                      
              </label>                                                                                                                        
              <div class="text-center"><span class="inner-label1">${imgHtml}</span></div>                                                     
                                                                                                                                              
              <label class="radio-label choice mt-4" style="display:none">                                                                    
                <input name="q${id}" type="radio" id="x${id}" value="0" checked>                                                              
                <span class="inner-label"></span>                                                                                             
              </label>                                                                                                                        
                                                                                                                                              
              ${opciones.map((op, i) => `                                                                                                     
                <label class="radio-label choice mt-4${i === 3 ? ' mb-2' : ''}">                                                              
                  <input name="q${id}" type="radio" id="${String.fromCharCode(97 + i)}${id}" value="${i + 1}">                                
                  <span class="inner-label">${op}</span>                                                                                      
                </label>                                                                                                                      
              `).join('')}                                                                                                                    
            </div>                                                                                                                            
          </div>                                                                                                                              
        `;                                                                                                                                    
      });                                                                                                                                     
                                                                                                                                              
      return [maxQ, correctAnswer, radioLists.join(""), pointValues];                                                                         
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
                                                                                                                                              
  function recordData(
  dni, nombre, cargo, empresa,
  tema, area, point, duracion,
  Ans, firmaBase64, capacitadorForm, duracionForm, detalle // 👈 nuevo parámetro
) {
  const ss = getSpreadsheetCapacitaciones();
  const hojaBD = ss.getSheetByName("B DATOS");
  const hojaMatriz = ss.getSheetByName("Matriz");
  const now = new Date();

  try {
    // === 1️⃣ Buscar datos del curso ===
    const numCursos = hojaMatriz.getLastColumn() - 4;
    const cursos = hojaMatriz.getRange(14, 5, 1, numCursos).getValues()[0];       // Fila 14 → Temas
    const puntajesMin = hojaMatriz.getRange(8, 5, 1, numCursos).getValues()[0];  // Fila 8 → Puntaje
    const temporalidades = hojaMatriz.getRange(6, 5, 1, numCursos).getValues()[0]; // Fila 6 → Vigencia
    const capacitadores = hojaMatriz.getRange(10, 5, 1, numCursos).getValues()[0]; // Fila 10 → Capacitador
    const horasLectivas = hojaMatriz.getRange(12, 5, 1, numCursos).getValues()[0]; // Fila 12 → Horas

    let puntajeMinimo = null;
    let temporalidad = null;
    let capacitador = "";
    let horasLectiva = "";

    for (let j = 0; j < cursos.length; j++) {
      if (String(cursos[j]).toLowerCase().trim() === String(tema).toLowerCase().trim()) {
        puntajeMinimo = parseFloat(puntajesMin[j]);
        temporalidad = parseInt(temporalidades[j]);
        capacitador = String(capacitadores[j] || "");
        horasLectiva = horasLectivas[j] || "";
        break;
      }
    }

    if (!capacitador && capacitadorForm) capacitador = capacitadorForm;
    if (!horasLectiva && duracionForm) horasLectiva = duracionForm;

    // === 2️⃣ Determinar estado ===
    let estadoFinal = "Pendiente";
    if (!isNaN(point) && point !== "-") {
      if (puntajeMinimo !== null) {
        estadoFinal = point >= puntajeMinimo ? "Aprobado" : "Reprobado";
      }
    }

    // === 3️⃣ Subir firma si existe ===
    let urlFirma = "";
    if (firmaBase64 && firmaBase64.startsWith("data:image/")) {
      const folder = DriveApp.getFolderById(foldefirmascap);
      const nombreArchivo = `${dni}_firma_${Date.now()}.png`;
      const blob = Utilities.newBlob(
        Utilities.base64Decode(firmaBase64.split(",")[1]),
        "image/png",
        nombreArchivo
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urlFirma = `https://lh5.googleusercontent.com/d/${file.getId()}`;
    }

    // === 4️⃣ Registrar datos ===
    const filaNueva = [
      "'" + dni,
      nombre,
      cargo,
      empresa,
      tema,
      area,
      point,
      now,
      horasLectiva || "",
      estadoFinal,
      temporalidad || "",
      capacitador || "",
      detalle || "", // 👈 Comentarios (columna M)
      urlFirma || ""
    ].concat(Ans || []);

    hojaBD.appendRow(filaNueva);

    return ["success", point, estadoFinal];
  } catch (error) {
    Logger.log("Error en recordData: " + error);
    return ["error", error.toString()];
  }
}
                                                                                                                                        
    //ESTRELLAS                                                                                                                               
    function guardarCalificacionEnFila(dni, calificacion) {                                                                                   
      const sheet = getSpreadsheetCapacitaciones().getSheetByName("B DATOS");                                                                 
      const lastRow = sheet.getLastRow();                                                                                                     
                                                                                                                                              
      // Solo leemos la columna A (donde está el DNI)                                                                                         
      const dniCol = sheet.getRange(1, 1, lastRow, 1).getValues(); // Columna A                                                               
                                                                                                                                              
      for (let i = lastRow - 1; i >= 0; i--) {                                                                                                
        if (dniCol[i][0] == dni) {                                                                                                            
          sheet.getRange(i + 1, 13).setValue(calificacion); // Columna M (13)                                                                 
          return true;                                                                                                                        
        }                                                                                                                                     
      }                                                                                                                                       
      return false;                                                                                                                           
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
    function getAllTopics() {                                                                                                                 
      const sheet = getSpreadsheetCapacitaciones().getSheetByName(quizData);                                                                  
      const data = sheet.getRange(2, 3, sheet.getLastRow()-1).getValues(); // Suponiendo que los temas están en la columna C                  
      const uniqueTopics = [...new Set(data.flat())].filter(String);                                                                          
      return uniqueTopics;                                                                                                                    
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
                                                                                                                                              
    //Configuracion matriz                                                                                                                    
function obtenerMatrizInvertida() {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("Matriz");
  if (!hoja) return null;

  const lastCol = hoja.getLastColumn();
  // leemos 13 filas horizontales (D2..D14: 2 nuevas + 10 datos + 1 índice)
  const rango = hoja.getRange(2, 4, 13, lastCol - 3).getValues(); // D2 en adelante, 13 filas

  const headers = rango.map(fila => fila[0]);
  const datos = rango.map(fila => fila.slice(1));
  const indice = datos.pop(); // la última fila del bloque será el índice (temas)

  const filas = datos[0].map((_, i) => {
    const id = indice[i];
    if (!id) return null;

    const fila = datos.map(f => f[i]);
    fila.unshift(id);

    return fila.map(celda =>
      celda instanceof Date
        ? Utilities.formatDate(celda, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
        : celda
    );
  }).filter(Boolean);

  headers.pop();
  headers.unshift("Curso");

  return { headers, filas };
}
                                                                                                                                     
                                                                                                                                              
function agregarRegistro(nuevaColumna) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("Matriz");
  const lastColumn = hoja.getLastColumn() + 1;

  // 13 valores: 1 nombre del curso + 12 propiedades (filas 2..13)
  if (nuevaColumna.length !== 13) throw new Error("Se requieren 13 valores: 1 curso + 12 datos");

  const datos = nuevaColumna.slice(1); // Las 12 propiedades (para filas 2..13)
  const curso = nuevaColumna[0];       // El nombre del curso

  // Escribe filas 2 a 13 en la nueva columna
  hoja.getRange(2, lastColumn, 12, 1).setValues(datos.map(d => [d]));

  // Escribe el índice del curso en la fila 14
  hoja.getRange(14, lastColumn).setValue(curso);
}
                                                                                                                                     
function actualizarRegistro(columnaEditada) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("Matriz");
  // los cursos están en la fila 14, desde columna E (col 5)
  const datos = hoja.getRange(14, 5, 1, hoja.getLastColumn() - 4).getValues()[0]; // fila 14, desde E

  const curso = columnaEditada[0];
  const colIndex = datos.indexOf(curso);

  if (colIndex === -1) throw new Error("Curso no encontrado");

  const datosNuevos = columnaEditada.slice(1); // sin el nombre del curso
  if (datosNuevos.length !== 12) throw new Error("Se requieren 12 valores para actualizar (filas 2..13)");

  const col = 5 + colIndex; // columna real donde está el curso

  // Actualiza filas 2 a 13 (12 filas)
  hoja.getRange(2, col, 12, 1).setValues(datosNuevos.map(d => [d]));

  // Actualiza el nombre del curso en fila 14
  hoja.getRange(14, col).setValue(curso);
}
                                                                                                                                
function eliminarRegistroPorCurso(nombreCurso) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("Matriz");
  if (!hoja) return;

  // D2 en adelante, altura = 13 filas (2..14)
  const rango = hoja.getRange(2, 4, 13, hoja.getLastColumn() - 3); // D2 en adelante
  const datos = rango.getValues();

  // Buscar la columna del índice (última fila del bloque = fila 14)
  const filaIndice = datos[datos.length - 1]; // esto apunta a la fila que contiene los nombres de curso (fila 14)
  const colIndex = filaIndice.indexOf(nombreCurso);

  if (colIndex === -1) return;

  // Borrar verticalmente todos los valores en esa columna (filas 2..14)
  for (let fila = 0; fila < datos.length; fila++) {
    hoja.getRange(fila + 2, 4 + colIndex).setValue(""); // columna D + colIndex
  }
}
                                                                                                                                  
    //EXAMEN                                                                                                                                  
    function obtenerMatrizExamenPaginado(offset, limit, filtro = "") {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(quizData);
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) {
    return { headers2: [], filas: [], total: 0 };
  }

  // Leer columnas A–K (11 columnas)
  const datos = hoja.getRange(1, 1, lastRow, 11).getValues();
  const [headers2, ...filas] = datos;

  let filtradas = filas;
  if (filtro) {
    const texto = filtro.toLowerCase();
    filtradas = filas.filter(fila =>
      fila.some(celda => String(celda).toLowerCase().includes(texto))
    );
  }

  const paginadas = filtradas.slice(offset, offset + limit);
  return {
    headers2,
    filas: paginadas,
    total: filtradas.length
  };
}

/**
 * 🔹 Recupera todas las preguntas pertenecientes a un tema específico.
 */
function obtenerPreguntasPorTema(tema) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(quizData);
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];

  const data = hoja.getRange(2, 1, lastRow - 1, 11).getValues();
  return data
    .filter(r => r[2] === tema)
    .map(r => ({
      id: r[0],
      pregunta: r[3],
      url: r[4],
      opciones: [r[5], r[6], r[7], r[8]],
      correcta: r[9],
      puntos: r[10]
    }));
}

/**
 * 🔹 Crea o actualiza múltiples preguntas de un mismo tema.
 * Si es edición: elimina físicamente las que ya no están en el modal.
 */
function guardarPreguntasMultiples(lista, idsOriginales, esEdicion) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(quizData);
  const lastRow = hoja.getLastRow();
  const dataExistente = lastRow > 1
    ? hoja.getRange(2, 1, lastRow - 1, 11).getValues()
    : [];

  if (esEdicion && idsOriginales?.length) {
    // 🔸 Identificar IDs que ya no existen en la nueva lista y eliminarlas físicamente
    const idsEliminar = idsOriginales.filter(id => !lista.some(p => p[0] === id));
    if (idsEliminar.length > 0) {
      // Buscar sus posiciones de fila (de abajo hacia arriba para no desajustar índices)
      const filasEliminar = [];
      dataExistente.forEach((r, i) => {
        if (idsEliminar.includes(r[0])) filasEliminar.push(i + 2);
      });
      filasEliminar.sort((a, b) => b - a).forEach(fila => hoja.deleteRow(fila));
    }
  }

  // 🔸 Actualizar o insertar cada pregunta
  lista.forEach(data => {
    if (!data[0]) {
      // Nueva pregunta → generar ID único
      const idUnico = "E" + Date.now().toString().slice(-7) + Math.floor(Math.random() * 100);
      data[0] = idUnico;
      hoja.appendRow(data);
    } else {
      // Buscar y actualizar si existe
      const index = dataExistente.findIndex(r => r[0] === data[0]);
      if (index >= 0) {
        hoja.getRange(index + 2, 1, 1, data.length).setValues([data]);
      } else {
        hoja.appendRow(data);
      }
    }
  });

  return true;
}

/**
 * 🔹 Elimina una pregunta por su ID.
 */
function eliminarPreguntaPorID(id) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(quizData);
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return false;

  const ids = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) {
      hoja.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}
                                                                                                                                
function obtenerOpcionesFormulario() {
  const ss = getSpreadsheetCapacitaciones();
  const hojaMatriz = ss.getSheetByName("Matriz");
  const hojaTemas = ss.getSheetByName("TEMAS");
  
  // --- 1. PROCESAR MATRIZ (Horizontal: Fila 11=Área, 14=Temas) ---
  const lastCol = hojaMatriz.getLastColumn();
  // Leemos desde la columna E (5) hasta el final
  const filaAreas = hojaMatriz.getRange(11, 5, 1, lastCol - 4).getValues()[0];
  const filaTemas = hojaMatriz.getRange(14, 5, 1, lastCol - 4).getValues()[0];

  let datosMatrizRelacion = [];
  let areasUnicasMatriz = [];

  filaAreas.forEach((area, i) => {
    const tema = filaTemas[i];
    if (area && tema) {
      const areaLimpia = area.toString().trim();
      const temaLimpio = tema.toString().trim();
      
      // Guardamos la relación para filtrar después
      datosMatrizRelacion.push([temaLimpio, areaLimpia]);
      
      // Lista para el primer select
      if (!areasUnicasMatriz.includes(areaLimpia)) {
        areasUnicasMatriz.push(areaLimpia);
      }
    }
  });

  // --- 2. PROCESAR HOJA TEMAS (Vertical) ---
  const lastRowTemas = hojaTemas.getLastRow();
  let datosTemasSheet = [];
  if (lastRowTemas > 1) {
    // Col B (Temas), Col C (Área)
    datosTemasSheet = hojaTemas.getRange(2, 2, lastRowTemas - 1, 2).getValues()
      .map(f => [f[0].toString().trim(), f[1].toString().trim()]);
  }

  return { 
    opciones1: areasUnicasMatriz.sort(), // Áreas de la Matriz
    datosMatrizSheet: datosMatrizRelacion, // Nueva relación [Tema, Área] de Matriz
    datosTemasSheet: datosTemasSheet       // Relación [Tema, Área] de hoja TEMAS
  };
}
                                                                                                                                     
    //MATRIZ CHECK                                                                                                                            
    function parseFecha(fechaStr) {                                                                                                           
      if (fechaStr instanceof Date) return fechaStr;                                                                                          
                                                                                                                                              
      if (typeof fechaStr === "string") {                                                                                                     
        const partes = fechaStr.split('/');                                                                                                   
        if (partes.length < 3) return new Date('');                                                                                           
        const [dia, mes, anioHora] = partes;                                                                                                  
        const [anio, hora] = anioHora.split(' ');                                                                                             
        return new Date(`${mes}/${dia}/${anio} ${hora || '00:00:00'}`);                                                                       
      }                                                                                                                                       
                                                                                                                                              
      return new Date('');                                                                                                                    
    }                                                                                                                                         
                                                                                                                                              
    function obtenerMatrizPermisos() {
  const datos = cargarDatosGlobales();
  const cursos = datos.cursos.filter(String);
  const cargos = datos.cargos.filter(String);
  const matriz = datos.matriz;
  const personalCargos = datos.personalCargos;
  const bdDatos = datos.bdDatos.slice(1);
  bdDatos.shift();

  const idxDNI = 0;             // Columna A
  const idxCargo = 2;           // Columna C
  const idxTema = 4;            // Columna E
  const idxPuntaje = 6;         // Columna G
  const idxFecha = 7;           // Columna H
  const idxHoras = 8;           // Columna I (Carga Horaria en minutos)
  const idxEstatus = 9;         // Columna J
  const idxTemporalidad = 10;   // Columna K

  const ahora = new Date();
  const mejoresPorDNIyCurso = {};

  bdDatos.forEach(row => {
    const dni = row[idxDNI];
    const curso = row[idxTema];
    const estatus = row[idxEstatus];
    const puntaje = parseFloat(row[idxPuntaje]) || 0;
    const fechaEval = parseFecha(row[idxFecha]);
    const dias = parseInt(row[idxTemporalidad]) || 0;

    if (estatus !== "Aprobado") return;

    const fechaLimite = new Date(fechaEval);
    fechaLimite.setDate(fechaEval.getDate() + dias);
    if (ahora > fechaLimite) return;

    const clave = `${dni}_${curso}`;
    if (!mejoresPorDNIyCurso[clave] || puntaje > (parseFloat(mejoresPorDNIyCurso[clave][idxPuntaje]) || 0)) {
      mejoresPorDNIyCurso[clave] = row;
    }
  });

  const datosFiltrados = Object.values(mejoresPorDNIyCurso);

  // === Cálculo de resumen por curso ===
  const resumen = cursos.map(curso => {
    const colIndex = cursos.indexOf(curso);
    const cargosAsignados = cargos.filter((_, i) => matriz[i][colIndex] === true);
    const personasAsignadas = personalCargos.filter(cargo => cargosAsignados.includes(cargo)).length;
    const personasAprobadas = datosFiltrados.filter(row =>
      cargosAsignados.includes(row[idxCargo]) && row[idxTema] === curso
    ).length;

    return { curso, asignados: personasAsignadas, aprobados: personasAprobadas };
  });

  // === Totales globales ===
  const totalProgramados = resumen.reduce((sum, item) => sum + item.asignados, 0);
  const totalAprobados = resumen.reduce((sum, item) => sum + item.aprobados, 0);

  // === Nuevos cálculos ===
  const totalCursosProgramados = cursos.length;
  const totalCursosRealizados = resumen.filter(r => r.aprobados > 0).length;
  const totalTrabajadores = new Set(bdDatos.map(r => r[idxDNI])).size;
  const totalHoras = datosFiltrados.reduce((sum, r) => sum + (parseFloat(r[idxHoras]) || 0), 0) / 60; // a horas

  return {
    cursos,
    cargos,
    matriz,
    resumen,
    totales: {
      cursosProgramados: totalCursosProgramados,
      cursosRealizados: totalCursosRealizados,
      trabajadores: totalTrabajadores,
      programados: totalProgramados,
      aprobados: totalAprobados,
      horas: totalHoras.toFixed(1) // redondear a 1 decimal
    }
  };
}


                                                                                                                                    
 function obtenerDetalleCurso(curso) {
  const ssCap = getSpreadsheetCapacitaciones();
  const hojaBDatos = ssCap.getSheetByName("B DATOS");
  const hojaMatriz = ssCap.getSheetByName("Matriz");
  const lastRow = hojaBDatos.getLastRow();
  const ahora = new Date();

  // 🔹 Leer solo columnas A–K (1–11) desde fila 2
  const datos = hojaBDatos.getRange(2, 1, lastRow - 1, 11).getValues();

  // 🔹 Leer cursos (fila 14), temporalidades (fila 6), certificación (fila 13)
  const cursos = hojaMatriz.getRange("E14:14").getValues()[0];
  const temporalidades = hojaMatriz.getRange("E6:6").getValues()[0];
  const certificaciones = hojaMatriz.getRange("E13:13").getValues()[0];

  // 🔹 Ubicar el índice del curso actual
  const colIndex = cursos.indexOf(curso);
  const tieneCertificacion = certificaciones[colIndex] === true || certificaciones[colIndex] === "VERDADERO";

  // Índices de columnas en B DATOS
  const idxDni = 0;             // Columna A
  const idxNombre = 1;          // Columna B
  const idxCargo = 2;           // Columna C
  const idxEmpresa = 3;         // Columna D
  const idxTema = 4;            // Columna E
  const idxPuntaje = 6;         // Columna G
  const idxFecha = 7;           // Columna H
  const idxEstatus = 9;         // Columna J
  const idxTemporalidad = 10;   // Columna K

  const mejoresPorDNI = {};

  // 🔹 Filtrar los mejores registros aprobados y vigentes
  datos.forEach(row => {
    if (row[idxTema] !== curso) return;
    if (row[idxEstatus] !== "Aprobado") return;

    const dni = row[idxDni];
    const puntaje = parseFloat(row[idxPuntaje]) || 0;
    const fechaEval = parseFecha(row[idxFecha]);
    const dias = parseInt(row[idxTemporalidad]) || 0;

    const fechaLimite = new Date(fechaEval);
    fechaLimite.setDate(fechaEval.getDate() + dias);
    if (ahora > fechaLimite) return;

    if (!mejoresPorDNI[dni] || puntaje > (parseFloat(mejoresPorDNI[dni][idxPuntaje]) || 0)) {
      mejoresPorDNI[dni] = row;
    }
  });

  // 🔹 Convertir los resultados en un arreglo de objetos formateados
  return Object.values(mejoresPorDNI).map(row => {
    let fechaCruda = row[idxFecha];
    let fechaFormateada = "";

    if (fechaCruda instanceof Date) {
      const dia = fechaCruda.getDate().toString().padStart(2, '0');
      const mes = (fechaCruda.getMonth() + 1).toString().padStart(2, '0');
      const anio = fechaCruda.getFullYear();
      fechaFormateada = `${dia}/${mes}/${anio}`;
    } else if (typeof fechaCruda === "string") {
      fechaFormateada = fechaCruda.split(" ")[0];
    }

    return {
      dni: row[idxDni],
      nombre: row[idxNombre],
      cargo: row[idxCargo],
      empresa: row[idxEmpresa],
      puntaje: row[idxPuntaje],
      estatus: row[idxEstatus],
      fecha: fechaFormateada,
      certificacion: tieneCertificacion ? "Con certificación" : "Curso sin certificación"
    };
  });
}
                                                                                        
                                                                                                                                              
    function actualizarCeldaCheckbox(fila, columna, valor) {                                                                                  
      const hoja = getSpreadsheetCapacitaciones().getSheetByName("Matriz");                                                                   
      hoja.getRange(fila, columna).setValue(valor);                                                                                           
    }                                                                                                                                         
                                                                                                                                              
    function actualizarAsignadosPorCurso(fila, columna, valor) {
  const ssCap = getSpreadsheetCapacitaciones();
  const hojaMatriz = ssCap.getSheetByName("Matriz");
  const hojaBDatos = ssCap.getSheetByName("B DATOS");
  const hojaPersonal = getSpreadsheetPersonal().getSheetByName("PERSONAL");

  // 🔹 Obtener todos los cursos (fila 14)
  const cursos = hojaMatriz.getRange("E14:14").getValues()[0];
  const lastRowMatriz = hojaMatriz.getLastRow();

  // 🔹 Calcular índice de columna dentro del rango de cursos
  const colIndex = columna - 5;
  if (colIndex < 0 || colIndex >= cursos.length) return;

  const curso = cursos[colIndex];

  // ✅ Actualizar el valor del checkbox
  hojaMatriz.getRange(fila, columna).setValue(valor);
  SpreadsheetApp.flush();

  // ✅ Leer la matriz desde fila 15 en adelante (cargos verticales)
  const cargosRange = hojaMatriz.getRange(15, 4, lastRowMatriz - 14, cursos.length + 1).getValues();
  const cargos = cargosRange.map(row => row[0]);
  const matriz = cargosRange.map(row => row.slice(1));

  // ✅ Cargos que tienen el curso marcado como asignado
  const cargosAsignados = new Set(
    cargos.filter((_, i) => matriz[i][colIndex] === true)
  );

  // ✅ Personal → Columna G contiene el cargo
  const personalCargos = hojaPersonal.getRange(2, 7, hojaPersonal.getLastRow() - 1).getValues().flat();
  const personasAsignadas = personalCargos.filter(cargo => cargosAsignados.has(cargo)).length;

  // ✅ Leer B DATOS (A–K)
  const lastRowBD = hojaBDatos.getLastRow();
  const datos = hojaBDatos.getRange(2, 1, lastRowBD - 1, 11).getValues();

  const ahora = new Date();
  const mejoresPorDNIyCurso = {};

  // ✅ Filtrar los mejores resultados por persona y curso
  for (const row of datos) {
    const [dni, , cargo, , tema, , puntajeStr, fechaStr, , estatus, diasStr] = row;

    if (tema !== curso || estatus !== "Aprobado") continue;

    const puntaje = parseFloat(puntajeStr) || 0;
    const fechaEval = parseFecha(fechaStr);
    const dias = parseInt(diasStr) || 0;

    const fechaLimite = new Date(fechaEval);
    fechaLimite.setDate(fechaEval.getDate() + dias);
    if (ahora > fechaLimite) continue;

    const clave = `${dni}_${tema}`;
    if (!mejoresPorDNIyCurso[clave] || puntaje > (parseFloat(mejoresPorDNIyCurso[clave][6]) || 0)) {
      mejoresPorDNIyCurso[clave] = row;
    }
  }

  const datosFiltrados = Object.values(mejoresPorDNIyCurso);

  // ✅ Contar aprobados entre los cargos asignados
  const personasAprobadas = datosFiltrados.filter(row =>
    cargosAsignados.has(row[2]) && row[4] === curso
  ).length;

  // ✅ Nueva salida con curso y valores actualizados
  return {
    curso,
    asignados: personasAsignadas,
    aprobados: personasAprobadas
  };
}
                                                                                                                               
                                                                           
    //BUSCADOR GENERAL                                                                                                                        
    function getDatosDesdeBDATOS(start = 0, length = 30, search = "", fechaDesde = "", fechaHasta = "") {                                     
      const sheetBDATOS = getSpreadsheetCapacitaciones().getSheetByName("B DATOS");                                                           
      const sheetMatriz = getSpreadsheetCapacitaciones().getSheetByName("Matriz");                                                            
      if (!sheetBDATOS || !sheetMatriz) throw new Error('Faltan hojas');                                                                      
                                                                                                                                              
      const data = sheetBDATOS.getRange("A1:M" + sheetBDATOS.getLastRow()).getValues();                                                       
      const headers = data[0];                                                                                                                
      headers.push("Aprobados / Previstos");                                                                                                  
      const hoy = new Date();                                                                                                                 
                                                                                                                                              
      const temasPorID = {};                                                                                                                  
      const cargoPorID = {};                                                                                                                  
                                                                                                                                              
      for (let i = 1; i < data.length; i++) {                                                                                                 
        const [id, , cargo, , tema, , , fecha, , estado, temporalidad] = data[i];                                                             
        if (!cargoPorID[id]) cargoPorID[id] = cargo;                                                                                          
                                                                                                                                              
        if (estado !== "Aprobado" || !(fecha instanceof Date) || isNaN(temporalidad)) continue;                                               
                                                                                                                                              
        const dias = Math.floor((hoy - fecha) / (1000 * 60 * 60 * 24));                                                                       
        if (dias >= 0 && dias <= parseInt(temporalidad)) {                                                                                    
          if (!temasPorID[id]) temasPorID[id] = new Set();                                                                                    
          temasPorID[id].add(tema);                                                                                                           
        }                                                                                                                                     
      }                                                                                                                                       
                                                                                                                                              
      const cargosMatriz = sheetMatriz.getRange("D15:D" + sheetMatriz.getLastRow()).getValues().flat().filter(c => c);
      const filaTemas = sheetMatriz.getRange("E14:14").getValues()[0];
      const cantidadTemas = filaTemas.filter(t => t).length;
      const matrizValores = sheetMatriz.getRange(15, 5, cargosMatriz.length, cantidadTemas).getValues();                                      
                                                                                                                                              
      const previstosPorCargo = {};                                                                                                           
      cargosMatriz.forEach((cargo, idx) => {                                                                                                  
        const fila = matrizValores[idx];                                                                                                      
        const verdaderos = fila.filter(val => val === true).length;                                                                           
        previstosPorCargo[cargo] = verdaderos;                                                                                                
      });                                                                                                                                     
                                                                                                                                              
      let filas = [];                                                                                                                         
                                                                                                                                              
      for (let i = 1; i < data.length; i++) {                                                                                                 
        const fila = [...data[i]];                                                                                                            
        const id = fila[0];                                                                                                                   
        const estado = fila[9];                                                                                                               
        const fecha = fila[7];                                                                                                                
                                                                                                                                              
        if (fecha instanceof Date) {                                                                                                          
          fila[7] = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");                                          
          fila._fechaOrden = fecha;                                                                                                           
        } else {                                                                                                                              
          fila._fechaOrden = new Date("1900-01-01");                                                                                          
        }                                                                                                                                     
                                                                                                                                              
        if (estado === "Aprobado") {                                                                                                          
          const aprobados = temasPorID[id] ? temasPorID[id].size : 0;                                                                         
          const cargo = cargoPorID[id] || "";                                                                                                 
          const previstos = previstosPorCargo[cargo] || 0;                                                                                    
          fila.push(`${aprobados}/${previstos}`);                                                                                             
          filas.push(fila);                                                                                                                   
        }                                                                                                                                     
      }                                                                                                                                       
                                                                                                                                              
      // Filtro por texto                                                                                                                     
      if (search) {                                                                                                                           
        const query = search.toLowerCase();                                                                                                   
        filas = filas.filter(row => row.some(cell => String(cell).toLowerCase().includes(query)));                                            
      }                                                                                                                                       
                                                                                                                                              
      // Filtro por fecha (columna 7 ya formateada)                                                                                           
      if (fechaDesde || fechaHasta) {                                                                                                         
        const desde = fechaDesde ? new Date(fechaDesde) : null;                                                                               
        const hasta = fechaHasta ? new Date(fechaHasta) : null;                                                                               
        filas = filas.filter(row => {                                                                                                         
          const f = row._fechaOrden;                                                                                                          
          return (!desde || f >= desde) && (!hasta || f <= hasta);                                                                            
        });                                                                                                                                   
      }                                                                                                                                       
                                                                                                                                              
      // Ordenar                                                                                                                              
      filas.sort((a, b) => b._fechaOrden - a._fechaOrden);                                                                                    
      filas = filas.map(row => {                                                                                                              
        delete row._fechaOrden;                                                                                                               
        return row;                                                                                                                           
      });                                                                                                                                     
                                                                                                                                              
      const totalFiltrado = filas.length;                                                                                                     
      const paginated = filas.slice(start, start + length);                                                                                   
                                                                                                                                              
      return {                                                                                                                                
        headers: headers,                                                                                                                     
        data: paginated,                                                                                                                      
        total: totalFiltrado                                                                                                                  
      };                                                                                                                                      
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
    function insertarYObtenerDatosCumplimiento(valor) {                                                                                       
      const hoja = getSpreadsheetCapacitaciones().getSheetByName("CUMPLIMIENTO🧍‍♂️");                                                          
      hoja.getRange("B3").setValue(valor);                                                                                                    
                                                                                                                                              
      const datosColA = hoja.getRange("A:A").getValues();                                                                                     
      let ultimaFilaValida = 0;                                                                                                               
                                                                                                                                              
      for (let i = datosColA.length - 1; i >= 0; i--) {                                                                                       
        const valorCelda = datosColA[i][0];                                                                                                   
        if (typeof valorCelda === "number" && valorCelda > 1) {                                                                               
          ultimaFilaValida = i + 1; // porque i es índice 0-based                                                                             
          break;                                                                                                                              
        }                                                                                                                                     
      }                                                                                                                                       
                                                                                                                                              
      if (ultimaFilaValida === 0) return []; // no hay filas válidas                                                                          
                                                                                                                                              
      const rango = hoja.getRange(`A1:H${ultimaFilaValida}`);                                                                                 
      return rango.getValues();                                                                                                               
    } 

                                                                                                                                             
    //CHARLAS                                                                                                                                                                                                                                                                    
// function saveData(value) {
//   const ws = getSpreadsheetCapacitaciones().getSheetByName("REGISTRO");
//   const randomId = Math.floor(Math.random() * 1e8);
//   const fila = [randomId].concat(value);
//   ws.appendRow(fila);
// }
    function saveData(value) {
  const ws = getSpreadsheetCapacitaciones().getSheetByName("REGISTRO");
  const randomId = Math.floor(Math.random() * 1e8);

  // Forzar campo "Comentarios" como texto (posición 11 si empieza desde 0)
  if (value[10] !== undefined && value[10] !== null) {
    value[10] = "'- " + value[10].toString();
  }

  const fila = [randomId].concat(value);
  ws.appendRow(fila);
}
                                                                                                                           
                                                                                                                                              
                                                                                                                                              
    function uploadFilesToDrive(files) {                                                                                                      
      var folder = DriveApp.getFolderById(foldercharlas); //ARCHIVO 1 "Charlas"                                                               
      var urls = files.map(function(file) {                                                                                                   
        var contentType = file.data.match(/^data:(.*?);/)[1];                                                                                 
        var base64Data = file.data.split(',')[1];                                                                                             
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, file.filename);                                         
        var createdFile = folder.createFile(blob);                                                                                            
        createdFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);                                                   
        return createdFile.getUrl();                                                                                                          
      });                                                                                                                                     
      return urls;                                                                                                                            
    }                                                                                                                                              
                                                                                                                                              
    //BUSCADOR CHARLAS                                                                                                                        
    function getDatosDesdeCharlas(start = 0, length = 30, search = "", fechaDesde = "", fechaHasta = "") {                                    
      const sheet = getSpreadsheetCapacitaciones().getSheetByName("REGISTRO");                                                                
      const lastRow = sheet.getLastRow();                                                                                                     
                                                                                                                                              
      // Solo columnas A–M → col 1 a 14                                                                                                       
      const data = sheet.getRange(1, 1, lastRow, 14).getValues();                                                                             
      const headers = data[0];                                                                                                                
      let rows = data.slice(1);                                                                                                               
                                                                                                                                              
      // Procesar fechas (columna B)                                                                                                          
      rows = rows.map(row => {                                                                                                                
        const fecha = row[1];                                                                                                                 
        if (fecha instanceof Date) {                                                                                                          
          row[1] = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");                                                    
          row._fechaOrden = fecha;                                                                                                            
        } else {                                                                                                                              
          row._fechaOrden = new Date("1900-01-01"); // Fecha inválida predeterminada                                                          
        }                                                                                                                                     
        return row;                                                                                                                           
      });                                                                                                                                     
                                                                                                                                              
      // Filtro por texto                                                                                                                     
      if (search) {                                                                                                                           
        const s = search.toLowerCase();                                                                                                       
        rows = rows.filter(row =>                                                                                                             
          row.some(cell => String(cell).toLowerCase().includes(s))                                                                            
        );                                                                                                                                    
      }                                                                                                                                       
                                                                                                                                              
      // Filtro por fecha (columna B)                                                                                                         
      if (fechaDesde || fechaHasta) {                                                                                                         
        const desdeDate = fechaDesde ? new Date(fechaDesde) : null;                                                                           
        const hastaDate = fechaHasta ? new Date(fechaHasta) : null;                                                                           
                                                                                                                                              
        rows = rows.filter(row => {                                                                                                           
          const fecha = row._fechaOrden;                                                                                                      
          if (desdeDate && fecha < desdeDate) return false;                                                                                   
          if (hastaDate && fecha > hastaDate) return false;                                                                                   
          return true;                                                                                                                        
        });                                                                                                                                   
      }                                                                                                                                       
                                                                                                                                              
      // Ordenar por fecha descendente                                                                                                        
      rows.sort((a, b) => b._fechaOrden - a._fechaOrden);                                                                                     
                                                                                                                                              
      // Eliminar campo auxiliar                                                                                                              
      rows = rows.map(row => {                                                                                                                
        delete row._fechaOrden;                                                                                                               
        return row;                                                                                                                           
      });                                                                                                                                     
                                                                                                                                              
      const totalFiltrado = rows.length;                                                                                                      
      const paginated = rows.slice(start, start + length);                                                                                    
                                                                                                                                              
      return {                                                                                                                                
        headers: headers,                                                                                                                     
        data: paginated,                                                                                                                      
        total: totalFiltrado                                                                                                                  
      };                                                                                                                                      
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
    //GUARDAR EDICION CHARLAS                                                                                                                 
    function actualizarFilaPorID(id, nuevosDatos) {                                                                                           
      const sheet = getSpreadsheetCapacitaciones().getSheetByName("REGISTRO");                                                                
      const lastRow = sheet.getLastRow();                                                                                                     
                                                                                                                                              
      // Solo columnas A–M → col 1 a 14                                                                                                       
      const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();                                                                         
                                                                                                                                              
      for (let i = 0; i < data.length; i++) {                                                                                                 
        if (String(data[i][0]).trim() === String(id).trim()) {                                                                                
          sheet.getRange(i + 2, 1, 1, nuevosDatos.length).setValues([nuevosDatos]);                                                           
          return true;                                                                                                                        
        }                                                                                                                                     
      }                                                                                                                                       
                                                                                                                                              
      throw new Error("ID no encontrado");                                                                                                    
    }                                                                                                                                         
                                                                                                                                              
                                                                                                                                              
    //ELIMINA CHARLA                                                                                                                          
    function eliminarRegistroPorId(id) {                                                                                                      
      const hoja = getSpreadsheetCapacitaciones().getSheetByName("REGISTRO");                                                                 
      const lastRow = hoja.getLastRow();                                                                                                      
                                                                                                                                              
      // Solo columnas A–M (1 a 14), sin encabezado                                                                                           
      const datos = hoja.getRange(2, 1, lastRow - 1, 14).getValues();                                                                         
                                                                                                                                              
      for (let i = 0; i < datos.length; i++) {                                                                                                
        if (String(datos[i][0]).trim() === String(id).trim()) {                                                                               
          hoja.deleteRow(i + 2); // +2 porque datos empieza en fila 2 y el índice i es base 0                                                 
          return `Registro con ID ${id} eliminado.`;                                                                                          
        }                                                                                                                                     
      }                                                                                                                                       
                                                                                                                                              
      throw new Error(`No se encontró el registro con ID ${id}`);                                                                             
    }                                                                                                                                         
   
//CREA CON BLOB
function generarCertificadoDesdeNombreYTema(nombre, tema) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName('CERTIFICADO');
  hoja.getRange('D9').setValue(nombre);
  hoja.getRange('D13').setValue(tema);

  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  const sheetId = hoja.getSheetId();
  const url = getSpreadsheetCapacitaciones().getUrl().replace(/edit$/, '');
  const exportUrl = url + 'export?format=pdf&gid=' + sheetId + '&range=A1:G24&portrait=false';

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + token }
  });

  const blob = response.getBlob().setName(`Certificado - ${nombre}.pdf`);
  const base64 = Utilities.base64Encode(blob.getBytes());

  return base64;
}


//IA CAPACITACIONES
//const API_KEY2 = "AIzaSyBlm8NHhMDagHHREeGrqLWRInfcHK6Y_bw";
function analizarArchivoConGemini(base64DataUrl) {
  //const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${API_KEY}`;
  
  const mimeType = base64DataUrl.match(/^data:(.*?);/)[1];
  const base64 = base64DataUrl.split(',')[1];

  // 🔁 Obtener las listas reales desde la hoja 'LISTAS'
  const listas = getTodasLasListas();

  // 🧠 Construir dinámicamente el prompt con esas listas
  const prompt = `
Extrae los siguientes campos desde el contenido del archivo adjunto.
Usa exactamente los valores disponibles en los siguientes menús:

Empresas válidas: ${listas.empresas.join(", ")}
Lugares válidos: ${listas.lugares.join(", ")}
Tipo de formación válidas: ${listas.capacitaciones.join(", ")}
Gestiones válidas: ${listas.areas.join(", ")}
Registrado por válidas: ${listas.trabajadores.join(", ")}

Devuelve los campos con este formato:

Fecha:  
Tema:  
Lugar:  
Tipo de formación:  
Capacitador:  
Empresa:  
Gestión:  
Duración (min):  
Asistentes:  
Comentarios:  
Registrado por:
`;

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        {
          inlineData: {
            mimeType: mimeType,
            data: base64
          }
        }
      ]
    }]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(geminiUrl, options);
  const json = JSON.parse(response.getContentText());

  if (json.candidates && json.candidates[0]?.content?.parts?.[0]?.text) {
    return json.candidates[0].content.parts[0].text;
  } else {
    return `⚠️ Error en respuesta de Gemini:\n${JSON.stringify(json)}`;
  }
}
/**
 * 🔹 Usa Gemini 2.5 Flash para generar preguntas con o sin archivo, con o sin texto adicional.
 * @param {string|null} base64DataUrl - Archivo en base64 (PDF, Word, etc) o null si no hay archivo.
 * @param {number} numPreguntas - Cantidad de preguntas a generar.
 * @param {string|null} textoBase - Texto de contexto o instrucciones del usuario.
 * @return {string} JSON con formato [{pregunta, respuestas[], correcta}]
 */
function analizarConGemini(base64DataUrl, numPreguntas, textoBase) {
  const hayArchivo = !!base64DataUrl;
  const hayTexto = !!textoBase && textoBase.trim() !== "";

  // 🔹 Construir el prompt dinámico según lo que el usuario proporcione
  let prompt = "";

  if (hayTexto && hayArchivo) {
    prompt = `
El usuario proporcionó un archivo y una instrucción adicional.
Analiza el documento adjunto teniendo en cuenta lo siguiente:
"""${textoBase}"""
Genera ${numPreguntas} preguntas de opción múltiple relevantes al contexto indicado.
`;
  } else if (hayArchivo) {
    prompt = `
Analiza el siguiente documento y genera ${numPreguntas} preguntas de opción múltiple.
`;
  } else if (hayTexto) {
    prompt = `
Genera ${numPreguntas} preguntas de opción múltiple basadas en el siguiente texto:
"""${textoBase}"""
`;
  } else {
    throw new Error("No se recibió ni texto ni archivo para procesar.");
  }

  prompt += `
Cada pregunta debe tener:
- 1 texto de pregunta.
- 4 posibles respuestas.
- 1 número que indique cuál es la correcta (1 a 4).

El formato de salida debe ser JSON puro, sin texto adicional, así:
[
  {
    "pregunta": "¿Cuál es el resultado de 2+2?",
    "respuestas": ["1","2","3","4"],
    "correcta": 4
  }
]
Devuelve **solo el JSON**, sin explicaciones ni texto fuera del arreglo.
`;

  // 🔹 Crear el payload dependiendo de si hay archivo o no
  const parts = [{ text: prompt }];
  if (hayArchivo) {
    const mimeType = base64DataUrl.match(/^data:(.*?);/)[1];
    const base64 = base64DataUrl.split(",")[1];
    parts.push({ inlineData: { mimeType, data: base64 } });
  }

  const payload = { contents: [{ parts }] };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;

  // 🔹 Llamada a Gemini
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  const texto = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";

  const inicio = texto.indexOf("[");
  const fin = texto.lastIndexOf("]");
  if (inicio === -1 || fin === -1) throw new Error("Respuesta inválida de Gemini.");

  // 🔹 Normalización del JSON
  try {
    const arr = JSON.parse(texto.substring(inicio, fin + 1));
    arr.forEach(p => {
      let c = p.correcta;
      if (typeof c === "string") {
        c = c.trim().toUpperCase();
        if (["A", "B", "C", "D"].includes(c)) c = ["A","B","C","D"].indexOf(c) + 1;
        else if (/^\d$/.test(c)) c = parseInt(c);
        else c = 1;
      }
      if (isNaN(c) || c < 1 || c > 4) c = 1;
      p.correcta = c;
    });
    return JSON.stringify(arr);
  } catch (e) {
    throw new Error("Error al interpretar el JSON generado por Gemini: " + e);
  }
}


///FORMULARIO DE CAPACITACIONES
// -----------------------------
// MÓDULO TEMAS (Server side)
// -----------------------------
var temasData = "TEMAS";

/**
 * Devuelve encabezados, filas paginadas y total (claves: headersTemas, filas, total).
 */
function obtenerMatrizTemasPaginado(offset, limit, filtro = "") {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(temasData);
  const lastRow = hoja.getLastRow();

  if (lastRow < 2) {
    return { headersTemas: [], filas: [], total: 0 };
  }

  // Leer encabezados desde la primera fila (todas las columnas presentes)
  const numCols = hoja.getLastColumn();
  const headersTemas = hoja.getRange(1, 1, 1, numCols).getDisplayValues()[0];

  // Leer solo datos desde la fila 2 hasta la última (sin encabezado)
  const numFilas = lastRow - 1;
  const datos = hoja.getRange(2, 1, numFilas, numCols).getDisplayValues();

  // Invertimos los datos para mostrar los últimos primero
  let filas = datos.reverse();

  // Filtro si se proporciona
  if (filtro) {
    const texto = String(filtro).toLowerCase();
    filas = filas.filter(fila =>
      fila.some(celda => String(celda).toLowerCase().includes(texto))
    );
  }

  // Paginación
  const paginadas = filas.slice(offset, offset + limit);

  return {
    headersTemas,
    filas: paginadas,
    total: filas.length
  };
}

/**
 * Agrega un nuevo registro en TEMAS.
 * data: objeto con campos: tema, area, capacitador, duracion, intentos, validez, estado, horaInicio, horaFin, examen, valoracion
 * Genera Código automático (prefijo "T" + 7 caracteres).
 */
function agregarTema(data) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(temasData);

  // Generar código único (T + 7 caracteres alfanum)
  const codigo = generarCodigoTema();
  const fila = [
    codigo,                         // Código
    data.tema || "",                // Temas
    data.area || "",                // Área
    data.capacitador || "",         // Capacitador
    data.duracion || "",            // Duración (Min)
    data.intentos || "",            // Intentos (Veces)
    data.validez || "",             // Validez (Min)
    data.estado || "Activo",        // Estado
    data.horaInicio || "",          // HoraInicio
    data.horaFin || "",             // HoraFin
    data.examen || "No",            // Examen
    data.valoracion || "No"         // Valoración
  ];

  hoja.appendRow(fila);
  return codigo;
}

/**
 * Actualiza un tema según su Código (data.codigo).
 * data debe contener las 12 columnas en formato de objeto (ver arriba).
 */
function actualizarTema(data) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(temasData);

  // Generar NUEVO código (como solicitaste)
  const nuevoCodigo = generarCodigoTema();
  const codigoBuscado = String(data.codigo || "").trim();

  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return false;

  const codigos = hoja.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < codigos.length; i++) {
    if (String(codigos[i][0]).trim() === codigoBuscado) {

      const fila = [
        nuevoCodigo,                // Nuevo código reemplaza al anterior
        data.tema || "",
        data.area || "",
        data.capacitador || "",
        data.duracion || "",
        data.intentos || "",
        data.validez || "",
        data.estado || "Activo",
        data.horaInicio || "",
        data.horaFin || "",
        data.examen || "No",
        data.valoracion || "No"
      ];

      hoja.getRange(i + 2, 1, 1, fila.length).setValues([fila]);

      // Retornar el nuevo código para el Swal
      return nuevoCodigo;
    }
  }
  return null;
}

/**
 * Elimina tema por Código.
 */
function eliminarTemaPorCodigo(codigo) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName(temasData);
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return false;

  const codigos = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
  const buscado = String(codigo).trim();
  for (let i = 0; i < codigos.length; i++) {
    if (String(codigos[i][0]).trim() === buscado) {
      hoja.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

/**
 * Genera código 'T' + 7 chars alfanum
 */
function generarCodigoTema() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let out = "T";
  for (let i = 0; i < 7; i++) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}

// ✅ Obtener tema por código (con caché)
function getTemaPorCodigo(codigo) {
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("TEMAS");
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return { error: "No hay datos en la hoja TEMAS." };

  // Leemos SOLO columnas necesarias: A–K (10 columnas)
  const datos = hoja.getRange(2, 1, ultimaFila - 1, 11).getValues();

  const codigoBuscado = String(codigo).trim().toUpperCase();
  const ahora = new Date();

  // Búsqueda desde abajo para encontrar último registro
  for (let i = datos.length - 1; i >= 0; i--) {
    const [
      codigoTema,  // A
      tema,        // B
      area,        // C
      capacitador, // D
      duracion,    // E
      intentos,    // F
      validez,     // G
      _,           // H (ya no se usa)
      fechaInicio, // I
      fechaFin,     // J
      status      //K
    ] = datos[i];

    // Optimización: normalizar **una sola vez**
    if (String(codigoTema).toUpperCase() !== codigoBuscado) continue;

    // Convertir fechas (de inmediato solo en coincidencia)
    const inicio = fechaInicio ? new Date(fechaInicio) : null;
    const fin = fechaFin ? new Date(fechaFin) : null;

    // Validación — deben existir ambas fechas
    if (!inicio || !fin) {
      return { error: "El curso no tiene fechas válidas configuradas." };
    }

    // Validación de rango
    if (ahora < inicio || ahora > fin) {
      return {
        error: `El curso no está disponible en este rango:
Inicio: ${inicio.toLocaleString()}
Fin: ${fin.toLocaleString()}`
      };
    }

    // Devolver datos válidos
    return { tema, area, capacitador, duracion, intentos, status };
  }

  return { error: "Código no encontrado." };
}



// ✅ Forzar actualización manual del caché
function actualizarCacheTemas() {
  const cache = CacheService.getScriptCache();
  cache.remove("TEMAS_CACHE");
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("TEMAS");
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return "Sin datos para actualizar";

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 8).getValues();
  cache.put("TEMAS_CACHE", JSON.stringify(datos), 300);
  return "Cache actualizado correctamente";
}

function getTemasDesdeBD() {  
  const cache = CacheService.getScriptCache();

  // === 1. Intentar obtener desde caché ===
  const cacheLista = cache.get("lista_temas");
  if (cacheLista) {
    try {
      return JSON.parse(cacheLista);  // Respuesta instantánea
    } catch (err) {}
  }

  // === 2. Leer desde hoja ===
  const hoja = getSpreadsheetCapacitaciones().getSheetByName("TEMAS");
  const lastRow = hoja.getLastRow();

  if (lastRow < 2) return [];

  // Leer solo columna B (temas)
  const valores = hoja.getRange(2, 2, lastRow - 1, 1).getValues();

  // Procesar
  const temasUnicos = [...new Set(valores.flat()
    .map(v => v && v.toString().trim())
    .filter(Boolean)
  )].sort();

  // === 3. Guardar en caché 10 min ===
  cache.put("lista_temas", JSON.stringify(temasUnicos), 600);

  return temasUnicos;
}


