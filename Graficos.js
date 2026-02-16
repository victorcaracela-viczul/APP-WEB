// Cache mejorado con gestión de memoria
let cacheGraficos = {
  datos: new Map(),
  spreadsheets: new Map(),
  timestamp: 0,
  maxSize: 50 // Límite de entradas en cacheGraficos
};

// Constantes
const CACHE_DURATION = 3 * 60 * 1000; // 3 minutos (reducido para mejor responsividad)
const SPREADSHEET_IDS = {
  graficos: "1J_v47ohrGj8S1XfWUdneH0l7mMTB8auSOEscHZwsM0g", // HOJA DE CALCULO GRAFICOS
  check: "12KkPwl_gfQCkqS9ZHsp4hS2fFkebgNbszvTDtZELObU", // HOJA DE CALCULO CHECK LIST 
  desvios: "1eIJfA7dAlkQ1rXcRGC2qSFnvZ-jYIPn8cA_TbUZcWZE" // HOJA DE CALCULO INSPECCIONES
};

// Gestión de cacheGraficos optimizada
function cleanCache() {
  if (cacheGraficos.datos.size > cacheGraficos.maxSize) {
    const entries = Array.from(cacheGraficos.datos.entries());
    entries.slice(0, Math.floor(cacheGraficos.maxSize / 2)).forEach(([key]) => {
      cacheGraficos.datos.delete(key);
    });
  }
}

function getCacheKey(tipo, params = '') {
  return `${tipo}_${params}`;
}

function setCache(key, dataGraficos) {
  cleanCache();
  cacheGraficos.datos.set(key, {
    dataGraficos: dataGraficos,
    timestamp: Date.now()
  });
}

function getCache(key) {
  const cached = cacheGraficos.datos.get(key);
  if (cached && (Date.now() - cached.timestamp) < CACHE_DURATION) {
    return cached.dataGraficos;
  }
  cacheGraficos.datos.delete(key);
  return null;
}

// Obtener spreadsheet con cacheGraficos mejorado
function getSpreadsheet(tipo) {
  if (!cacheGraficos.spreadsheets.has(tipo)) {
    try {
      cacheGraficos.spreadsheets.set(tipo, SpreadsheetApp.openById(SPREADSHEET_IDS[tipo]));
    } catch (error) {
      console.error(`Error abriendo spreadsheet ${tipo}:`, error);
      throw new Error(`No se pudo acceder al spreadsheet ${tipo}`);
    }
  }
  return cacheGraficos.spreadsheets.get(tipo);
}

// Función para obtener datos con optimización de rangos
function obtenerDatosOptimizados() {
  const cacheKey = getCacheKey('datos_principales');
  const cached = getCache(cacheKey);
  if (cached) {
    console.log("Usando datos desde cacheGraficos");
    return cached;
  }

  console.log("Cargando datos optimizados...");
  const startTime = Date.now();
  
  try {
    // Obtener hojas
    const hojaGraficos = getSpreadsheet('graficos').getSheetByName("B DATOS");
    const hojaDesvios = getSpreadsheet('desvios').getSheetByName("B DATOS");
    const hojaCheck = getSpreadsheet('check').getSheetByName("B DATOS");
    
    // Obtener solo rangos necesarios de una vez
    const [principales, empresas] = obtenerDatosPrincipales(hojaGraficos);
    const desvios = obtenerDatosDesvios(hojaDesvios);
    const observaciones = obtenerDatosObservaciones(hojaCheck);
    
    const resultado = {
      principales,
      empresas,
      desvios,
      observaciones,
      timestamp: Date.now()
    };
    
    setCache(cacheKey, resultado);
    console.log(`Datos cargados en ${Date.now() - startTime}ms`);
    return resultado;
    
  } catch (error) {
    console.error("Error cargando datos:", error);
    throw error;
  }
}

// Función optimizada para datos principales
function obtenerDatosPrincipales(hoja) {
  const lastRow = hoja.getLastRow();
  if (lastRow <= 1) return [[], []];
  
  // Obtener solo las columnas necesarias (B2:T)
  const rango = hoja.getRange(2, 2, lastRow - 1, 19);
  const valores = rango.getValues();
  
  const principales = [];
  const empresasSet = new Set();
  
  // Procesar desde el final hacia adelante (más eficiente para datos recientes)
  for (let i = valores.length - 1; i >= 0; i--) {
    const fila = valores[i];
    
    // Validación rápida
    if (!fila[0]) continue;
    
    const dato = {
      periodo: fila[0],
      empresa: fila[1] || '',
      trabajadores: Number(fila[2]) || 0,
      ic: Number(fila[3]) || 0,
      desvios: Number(fila[4]) || 0,
      incidentes: Number(fila[5]) || 0,
      aLeve: Number(fila[6]) || 0,
      aGrave: Number(fila[7]) || 0,
      aFatal: Number(fila[8]) || 0,
      diasPerdidos: Number(fila[9]) || 0,
      enfermedadesOcupacionales: Number(fila[10]) || 0,
      hht: Number(fila[11]) || 0,
      metaIC: Number(fila[12]) || 0,
      metaIF: Number(fila[13]) || 0,
      if: Number(fila[14]) || 0,
      metaIG: Number(fila[15]) || 0,
      ig: Number(fila[16]) || 0,
      metaIA: Number(fila[17]) || 0,
      ia: Number(fila[18]) || 0
    };
    
    principales.push(dato);
    if (dato.empresa) empresasSet.add(dato.empresa);
  }
  
  // Revertir para orden cronológico
  principales.reverse();
  
  return [principales, Array.from(empresasSet).sort()];
}

// Función optimizada para desvíos
function obtenerDatosDesvios(hoja) {
  const lastRow = hoja.getLastRow();
  if (lastRow <= 1) return [];
  
  // Obtener encabezados una sola vez
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const indices = {
    fecha: encabezados.findIndex(h => h && h.toString().toLowerCase().includes('fecha modificación')),
    empresa: encabezados.findIndex(h => h && h.toString().toLowerCase().includes('empresa')),
    estado: encabezados.findIndex(h => h && h.toString().toLowerCase().includes('estado')),
    potencial: encabezados.findIndex(h => h && h.toString().toLowerCase().includes('potencial')),
    reportante: encabezados.findIndex(h => h && h.toString().toLowerCase().includes('reportante')),
    tipo: encabezados.findIndex(h => h && h.toString().toLowerCase().includes('tipo'))
  };
  
  // Verificar que tenemos las columnas mínimas
  if (indices.fecha < 0 || indices.empresa < 0) {
    console.warn("No se encontraron columnas necesarias en desvíos");
    return [];
  }
  
  // Obtener solo las columnas necesarias
  const columnasNecesarias = Math.max(...Object.values(indices).filter(i => i >= 0)) + 1;
  const rango = hoja.getRange(2, 1, lastRow - 1, columnasNecesarias);
  const valores = rango.getValues();
  
  const desvios = [];
  
  // Filtrar datos válidos en una sola pasada
  for (let i = valores.length - 1; i >= 0; i--) {
    const fila = valores[i];
    
    // Validación de fecha
    const fecha = fila[indices.fecha];
    if (!fecha || !(fecha instanceof Date)) continue;
    
    const desvio = {
      fecha: fecha,
      empresa: fila[indices.empresa] || '',
      estado: fila[indices.estado] || '',
      potencial: fila[indices.potencial] || '',
      reportante: fila[indices.reportante] || '',
      tipo: fila[indices.tipo] || ''
    };
    
    desvios.push(desvio);
  }
  
  desvios.reverse();
  return desvios;
}

// Función optimizada para observaciones
function obtenerDatosObservaciones(hoja) {
  const lastRow = hoja.getLastRow();
  if (lastRow <= 1) return [];
  
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const indices = {
    fecha: encabezados.findIndex(h => h && h.toString().toLowerCase().includes('fecha')),
    empresa: encabezados.findIndex(h => h && h.toString().toLowerCase().includes('empresa')),
    estado: encabezados.findIndex(h => h && h.toString().toLowerCase().includes('estado'))
  };
  
  if (indices.fecha < 0 || indices.empresa < 0) {
    console.warn("No se encontraron columnas necesarias en observaciones");
    return [];
  }
  
  const columnasNecesarias = Math.max(...Object.values(indices).filter(i => i >= 0)) + 1;
  const rango = hoja.getRange(2, 1, lastRow - 1, columnasNecesarias);
  const valores = rango.getValues();
  
  const observaciones = [];
  
  for (let i = valores.length - 1; i >= 0; i--) {
    const fila = valores[i];
    
    const fecha = fila[indices.fecha];
    if (!fecha || !(fecha instanceof Date)) continue;
    
    const observacion = {
      fecha: fecha,
      empresa: fila[indices.empresa] || '',
      estado: fila[indices.estado] || ''
    };
    
    observaciones.push(observacion);
  }
  
  observaciones.reverse();
  return observaciones;
}

// Funciones principales (optimizadas)
function obtenerDatosNiveles(fechaInicioGraficos, fechaFinGraficos, empresa) {
  const datos = obtenerDatosOptimizados();
  return procesarDatosNiveles(datos.principales, fechaInicioGraficos, fechaFinGraficos, empresa);
}

function obtenerDatosMensuales(fechaFinGraficos, empresa) {
  const datos = obtenerDatosOptimizados();
  return procesarDatosMensuales(datos.principales, fechaFinGraficos, empresa);
}

function obtenerEmpresasUnicas() {
  const datos = obtenerDatosOptimizados();
  return datos.empresas;
}

function obtenerResumenDesviosPorMes(fechaInicioGraficos, fechaFinGraficos, empresa) {
  const datos = obtenerDatosOptimizados();
  return procesarResumenDesvios(datos.desvios, fechaInicioGraficos, fechaFinGraficos, empresa);
}

function obtenerResumenEstadosPorMes(fechaInicioGraficos, fechaFinGraficos, empresa) {
  const datos = obtenerDatosOptimizados();
  return procesarResumenEstados(datos.observaciones, fechaInicioGraficos, fechaFinGraficos, empresa);
}

// Procesadores optimizados
function procesarDatosNiveles(principales, fechaInicioGraficos, fechaFinGraficos, empresa) {
  const inicio = new Date(fechaInicioGraficos);
  const fin = new Date(fechaFinGraficos);
  
  const acumulado = [0, 0, 0, 0, 0, 0, 0];
  const originales = [0, 0, 0, 0, 0, 0];
  
  let contadorFilas = 0;
  let sumaMetaIC = 0;
  let contadorMetaIC = 0;
  
  // Filtrar y procesar en una sola pasada
  const filasFiltradas = principales.filter(fila => {
    const fecha = new Date(fila.periodo);
    return (!empresa || fila.empresa === empresa) && 
           !isNaN(fecha.getTime()) && fecha >= inicio && fecha <= fin;
  });
  
  filasFiltradas.forEach(fila => {
    contadorFilas++;
    
    // Acumular valores
    acumulado[0] += fila.trabajadores;
    acumulado[2] += fila.desvios;
    acumulado[3] += fila.incidentes;
    acumulado[4] += fila.aLeve;
    acumulado[5] += fila.aGrave;
    acumulado[6] += fila.aFatal;
    
    // IC calculado
    if (fila.trabajadores && fila.ic && fila.metaIC && fila.metaIC !== 0) {
      acumulado[1] += (fila.trabajadores / fila.metaIC) * fila.ic;
    }
    
    if (fila.metaIC) {
      sumaMetaIC += fila.metaIC;
      contadorMetaIC++;
    }
    
    // Originales (acumulados)
    originales[0] += fila.ic;
    originales[1] += fila.desvios;
    originales[2] += fila.incidentes;
    originales[3] += fila.aLeve;
    originales[4] += fila.aGrave;
    originales[5] += fila.aFatal;
  });
  
  // Transformaciones para pirámide (constantes optimizadas)
  const factores = [1, 1, 1, 0.065, 0.029, 0.0045, 0.005];
  for (let i = 3; i < 7; i++) {
    acumulado[i] /= factores[i];
  }
  
  // Promedio solo para IC
  if (contadorFilas > 0) {
    originales[0] /= contadorFilas;
  }
  
  const promedioMetaIC = contadorMetaIC > 0 ? sumaMetaIC / contadorMetaIC : 0;
  
  return {
    etiquetas: ['Trabajadores', 'IC', 'Desvíos', 'Incidentes', 'A. Leve', 'A. Grave', 'A. Fatal'],
    datos: acumulado,
    valoresOriginales: originales,
    promedioMetaIC
  };
}

function procesarDatosMensuales(principales, fechaFinGraficos, empresa) {
  const fechaFinObj = new Date(fechaFinGraficos);
  const fechaInicioGraficos = new Date(fechaFinObj.getFullYear() - 1, fechaFinObj.getMonth() + 1, 1);
  
  // Pre-calcular fechas de meses
  const rangosMeses = [];
  const etiquetasMeses = [];
  
  for (let i = 0; i < 12; i++) {
    const fechaMes = new Date(fechaInicioGraficos.getFullYear(), fechaInicioGraficos.getMonth() + i, 1);
    const siguienteMes = new Date(fechaMes.getFullYear(), fechaMes.getMonth() + 1, 1);
    
    rangosMeses.push({ inicio: fechaMes, fin: siguienteMes });
    etiquetasMeses.push(fechaMes.toLocaleDateString('es-ES', { month: 'short', year: '2-digit' }));
  }
  
  // Filtrar datos una sola vez
  const datosFiltrados = principales.filter(fila => {
    const fecha = new Date(fila.periodo);
    return (!empresa || fila.empresa === empresa) && 
           !isNaN(fecha.getTime()) && fecha >= fechaInicioGraficos && fecha < rangosMeses[11].fin;
  });
  
  // Inicializar resultado
  const campos = [
    'trabajadores', 'ic', 'metaIC', 'desvios', 'incidentes', 'aLeve', 'aGrave', 'aFatal',
    'enfermedadesOcupacionales', 'diasPerdidos', 'hht', 'if', 'metaIF', 'ig', 'metaIG', 'ia', 'metaIA'
  ];
  
  const resultado = { meses: etiquetasMeses };
  campos.forEach(campo => resultado[campo] = []);
  resultado.ifCalculado = [];
  resultado.igCalculado = [];
  resultado.iaCalculado = [];
  
  // Procesar cada mes
  rangosMeses.forEach((rango, mes) => {
    // Datos del mes actual
    const datosMes = datosFiltrados.filter(fila => {
      const fecha = new Date(fila.periodo);
      return fecha >= rango.inicio && fecha < rango.fin;
    });
    
    if (datosMes.length > 0) {
      // Calcular sumas y promedios
      const sumas = datosMes.reduce((acc, fila) => {
        campos.forEach(campo => {
          acc[campo] += fila[campo] || 0;
        });
        return acc;
      }, Object.fromEntries(campos.map(campo => [campo, 0])));
      
      const cant = datosMes.length;
      
      // Promedios para métricas de gestión
      ['trabajadores', 'ic', 'metaIC', 'if', 'metaIF', 'ig', 'metaIG', 'ia', 'metaIA'].forEach(campo => {
        resultado[campo].push(sumas[campo] / cant);
      });
      
      // Sumas para eventos
      ['desvios', 'incidentes', 'aLeve', 'aGrave', 'aFatal', 'enfermedadesOcupacionales', 'diasPerdidos', 'hht'].forEach(campo => {
        resultado[campo].push(sumas[campo]);
      });
      
    } else {
      // Llenar con ceros
      campos.forEach(campo => resultado[campo].push(0));
    }
    
    // Calcular índices acumulados (12 meses móviles)
    const fechaInicioAcum = new Date(rango.inicio.getFullYear() - 1, rango.inicio.getMonth() + 1, 1);
    const datosAcum = datosFiltrados.filter(fila => {
      const fecha = new Date(fila.periodo);
      return fecha >= fechaInicioAcum && fecha < rango.fin;
    });
    
    const acum = datosAcum.reduce((acc, fila) => {
      acc.aGrave += fila.aGrave || 0;
      acc.aFatal += fila.aFatal || 0;
      acc.diasPerdidos += fila.diasPerdidos || 0;
      acc.hht += fila.hht || 0;
      return acc;
    }, { aGrave: 0, aFatal: 0, diasPerdidos: 0, hht: 0 });
    
    // Índices calculados
    const ifCalc = acum.hht > 0 ? ((acum.aGrave + acum.aFatal) * 1000000) / acum.hht : 0;
    const igCalc = acum.hht > 0 ? (acum.diasPerdidos * 1000000) / acum.hht : 0;
    const iaCalc = (ifCalc * igCalc) / 1000;
    
    resultado.ifCalculado.push(ifCalc);
    resultado.igCalculado.push(igCalc);
    resultado.iaCalculado.push(iaCalc);
  });
  
  return resultado;
}

function procesarResumenDesvios(desvios, fechaInicioGraficos, fechaFinGraficos, empresa) {
  const inicio = new Date(fechaInicioGraficos);
  const fin = new Date(fechaFinGraficos);
  const resumen = {};
  
  // Filtrar y agrupar en una sola pasada
  desvios.forEach(desvio => {
    const fecha = desvio.fecha;
    if (!(fecha instanceof Date) || fecha < inicio || fecha > fin) return;
    if (empresa && desvio.empresa !== empresa) return;
    
    const mesClave = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM");
    
    if (!resumen[mesClave]) {
      resumen[mesClave] = {
        potenciales: {
          Abierto: { Alto: 0, Medio: 0, Bajo: 0 },
          Cerrado: { Alto: 0, Medio: 0, Bajo: 0 }
        },
        reportantes: new Set(),
        tipo: { Acto: 0, Condicion: 0 }
      };
    }
    
    const estado = desvio.estado || "";
    const potencial = desvio.potencial || "";
    const reportante = desvio.reportante || "";
    const tipo = desvio.tipo || "";
    
    // Acumular por estado y potencial
    if (estado in resumen[mesClave].potenciales && 
        potencial in resumen[mesClave].potenciales[estado]) {
      resumen[mesClave].potenciales[estado][potencial]++;
    }
    
    // Reportantes únicos
    if (reportante) resumen[mesClave].reportantes.add(reportante);
    
    // Tipo (usando regex más eficiente)
    const tipoLower = tipo.toLowerCase();
    if (tipoLower.includes('acto')) resumen[mesClave].tipo.Acto++;
    if (tipoLower.includes('condición') || tipoLower.includes('condicion')) resumen[mesClave].tipo.Condicion++;
  });
  
  // Convertir a array ordenado
  return Object.keys(resumen).sort().map(mes => ({
    mes,
    abiertos: resumen[mes].potenciales.Abierto,
    cerrados: resumen[mes].potenciales.Cerrado,
    reportantesUnicos: resumen[mes].reportantes.size,
    tipo: resumen[mes].tipo
  }));
}

function procesarResumenEstados(observaciones, fechaInicioGraficos, fechaFinGraficos, empresa) {
  const inicio = new Date(fechaInicioGraficos);
  const fin = new Date(fechaFinGraficos);
  const resumen = {};
  
  // Estados válidos predefinidos
  const estadosValidos = ['Conforme', 'Abierto', 'Cerrado'];
  
  observaciones.forEach(obs => {
    const fecha = obs.fecha;
    if (!(fecha instanceof Date) || fecha < inicio || fecha > fin) return;
    if (empresa && obs.empresa !== empresa) return;
    
    const mesClave = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM");
    const estado = obs.estado || "";
    
    if (!resumen[mesClave]) {
      resumen[mesClave] = { Conforme: 0, Abierto: 0, Cerrado: 0 };
    }
    
    if (estadosValidos.includes(estado)) {
      resumen[mesClave][estado]++;
    }
  });
  
  return Object.keys(resumen).sort().map(mes => ({
    mes,
    Conforme: resumen[mes].Conforme,
    Abierto: resumen[mes].Abierto,
    Cerrado: resumen[mes].Cerrado
  }));
}

// Funciones de utilidad para mantenimiento
function limpiarCache() {
  cacheGraficos.datos.clear();
  cacheGraficos.timestamp = 0;
  console.log("Cache limpiado manualmente");
  return "Cache limpiado correctamente";
}

function obtenerEstadoCache() {
  return {
    entradas: cacheGraficos.datos.size,
    ultimaActualizacion: new Date(cacheGraficos.timestamp).toLocaleString(),
    spreadsheetsCargados: Array.from(cacheGraficos.spreadsheets.keys()),
    memoryUsage: `${cacheGraficos.datos.size}/${cacheGraficos.maxSize} entradas`
  };
}

// Función para pre-cargar datos (útil para warming up)
function precargarDatos() {
  try {
    const datos = obtenerDatosOptimizados();
    return {
      success: true,
      message: "Datos precargados correctamente",
      stats: {
        principales: datos.principales.length,
        empresas: datos.empresas.length,
        desvios: datos.desvios.length,
        observaciones: datos.observaciones.length
      }
    };
  } catch (error) {
    return {
      success: false,
      message: "Error al precargar datos: " + error.message
    };
  }
}

// Función de diagnóstico
function diagnosticarRendimiento() {
  const start = Date.now();
  
  try {
    // Test de carga de datos
    const datos = obtenerDatosOptimizados();
    const loadTime = Date.now() - start;
    
    // Test de procesamiento
    const processStart = Date.now();
    const empresas = datos.empresas;
    const hoy = new Date();
    const hace30dias = new Date(hoy.getTime() - 30 * 24 * 60 * 60 * 1000);
    
    procesarDatosNiveles(datos.principales, hace30dias.toISOString().split('T')[0], 
                        hoy.toISOString().split('T')[0], empresas[0] || '');
    
    const processTime = Date.now() - processStart;
    
    return {
      loadTime: `${loadTime}ms`,
      processTime: `${processTime}ms`,
      totalTime: `${Date.now() - start}ms`,
      cacheStatus: obtenerEstadoCache(),
      dataStats: {
        principales: datos.principales.length,
        empresas: datos.empresas.length,
        desvios: datos.desvios.length,
        observaciones: datos.observaciones.length
      }
    };
    
  } catch (error) {
    return {
      error: error.message,
      loadTime: `${Date.now() - start}ms (falló)`
    };
  }
}



//PRONOSTICO

const CONFIG = {
  CACHE_DAYS: 7,
  ANALYSIS_DAYS: 14,
  SHEETS: {
    IA: "IA",
    DATA: "B DATOS"
  },
  COLUMNS: {
    LOCATION: 4,
    TYPE: 7,
    DESCRIPTION: 8,
    DATE: 10
  }
};

/**
 * Función principal optimizada para generar pronósticos
 */
function generarPronosticoConGemini(fechaInicioStr, fechaFinStr) {
  try {
    const cacheGraficos = new PronosticoCache();
    
    // Verificar caché primero
    const cachedResult = cacheGraficos.get();
    if (cachedResult) {
      return cachedResult;
    }
    
    // Generar nuevo pronóstico
    const analyzer = new DataAnalyzer();
    const generator = new PronosticoGenerator();
    
    const dataGraficos = analyzer.getFilteredData(fechaInicioStr, fechaFinStr);
    if (!dataGraficos.length) {
      return "📊 No hay suficientes datos en el período analizado para generar un pronóstico confiable.";
    }
    
    const resultado = generator.generate(dataGraficos);
    
    // Guardar en caché
    cacheGraficos.set(resultado);
    
    return resultado;
    
  } catch (error) {
    console.error('Error en generarPronosticoConGemini:', error);
    return `⚠️ Error al generar pronóstico: ${error.message}`;
  }
}

/**
 * Clase para manejo inteligente de caché
 */
class PronosticoCache {
  constructor() {
    this.sheet = getDesviosSpreadsheet().getSheetByName(CONFIG.SHEETS.IA);
    this.today = new Date();
  }
  
  get() {
    const lastRow = this.sheet.getLastRow();
    if (lastRow <= 1) return null;
    
    const [fecha, resultado] = this.sheet.getRange(lastRow, 1, 1, 2).getValues()[0];
    const daysDiff = Math.floor((this.today - new Date(fecha)) / (1000 * 60 * 60 * 24));
    
    return daysDiff < CONFIG.CACHE_DAYS ? resultado : null;
  }
  
  set(resultado) {
    this.sheet.appendRow([this.today, resultado]);
    this.cleanup();
  }
  
  cleanup() {
    const lastRow = this.sheet.getLastRow();
    if (lastRow > 50) { // Mantener solo últimos 50 registros
      this.sheet.deleteRows(2, lastRow - 50);
    }
  }
}

/**
 * Clase para análisis optimizado de datos
 */
class DataAnalyzer {
  constructor() {
    this.sheet = getDesviosSpreadsheet().getSheetByName(CONFIG.SHEETS.DATA);
  }
  
  getFilteredData(fechaInicioStr, fechaFinStr) {
    const dataGraficos = this.sheet.getRange(2, 1, this.sheet.getLastRow() - 1, 22).getValues();
    const fechaInicioGraficos = new Date(fechaInicioStr);
    const fechaFinGraficos = new Date(fechaFinStr);
    
    return dataGraficos.filter(row => this.isValidRow(row, fechaInicioGraficos, fechaFinGraficos))
               .map(row => this.formatRowData(row));
  }
  
  isValidRow(row, fechaInicioGraficos, fechaFinGraficos) {
    const fechaTexto = row[CONFIG.COLUMNS.DATE];
    if (!fechaTexto || typeof fechaTexto !== "string") return false;
    
    const fecha = new Date(fechaTexto.replace(/'/g, ''));
    return fecha >= fechaInicioGraficos && fecha <= fechaFinGraficos;
  }
  
  formatRowData(row) {
    return {
      lugar: row[CONFIG.COLUMNS.LOCATION] || "Ubicación no especificada",
      descripcion: row[CONFIG.COLUMNS.DESCRIPTION] || "Sin descripción disponible",
      tipo: row[CONFIG.COLUMNS.TYPE] || "Tipo no clasificado"
    };
  }
}

/**
 * Clase para generación de pronósticos con IA
 */
class PronosticoGenerator {
  generate(data) {
    const prompt = this.buildPrompt(data);
    return this.callGeminiAPI(prompt);
  }
  
  buildPrompt(data) {
    const incidents = data.map(d => 
      `• ${d.lugar}: ${d.descripcion} (${d.tipo})`
    ).join('\n');
    
    return `Como experto en seguridad industrial y minera, analiza estos ${data.length} actos inseguros o condiciones subestandares recientes y genera un pronóstico de riesgo semanal profesional y específico:

DATOS HISTÓRICOS:
${incidents}

FORMATO REQUERIDO:
"Se han registrado ${data.length} desvíos (condiciones subestándar y/o actos inseguros).
La zona o zonas es [UBICACIÓN ESPECÍFICA] con un [PORCENTAJE]% de probabilidad de que ocurra un incidente o accidente.

✅3️⃣ principales factores críticos identificados: [3 PATRONES PRINCIPALES DETECTADOS, MUY CLAROS Y DIRECTOS], no mas de 250 caracteres,
✅3️⃣ principales recomendaciones preventivas: [3 RECOMENDACIONES PREVENTIVAS ESPECÍFICAS Y PRÁCTICAS], no mas de 250 caracteres"`;
  }
  
  callGeminiAPI(prompt) {
    const payload = {
      contents: [{ parts: [{ text: prompt }] }]
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    };
    
    const response = UrlFetchApp.fetch(geminiUrl, options);
    const result = JSON.parse(response.getContentText());
    
    return result?.candidates?.[0]?.content?.parts?.[0]?.text || 
           "🤖 No se pudo generar el pronóstico. Intente nuevamente.";
  }
}