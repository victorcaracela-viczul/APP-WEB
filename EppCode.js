
let cachedEPP = null;
function getSpreadsheetEPP() {
  if (!cachedEPP) {
    cachedEPP = SpreadsheetApp.openById("1Mxy5SkDdLy1Ihct844uLq5ZALe-RFarDfWo9j65kBcE");
  }
  return cachedEPP;
}
 const FOLDER_IDEPP = '1rzdA2M0abUuT83i9YFF7-W8eLHWrs84s'; //CARPETA EPP - FIRMA
/** ================== CONFIG / SHEETS ================== **/
const SHEPP = {
  STOCK:     'STOCK',
  MOV:       'MOVIMIENTOS',
  REGISTRO:  'REGISTRO',
  MATRIZ:    'MATRIZ',
  ALMACENES: 'ALMACENES',
  //PERSONAL:  'PERSONAL'
};


// ======== COLUMN INDEX MAP (1-based) — audit from your Excel ========
// (Si cambias el orden de columnas en la hoja, ACTUALIZA solo estos índices.)
const IDX = {
  STOCK: { ID:1, ALMACEN:2, PRODUCTO:3, VARIANTE:4, CATEGORIA:5, STOCK:6, STOCK_MINIMO:7, PRECIO:8, IMG:9, FILE:10 },
  MOV:   { ID_MOV:1, FECHA:2, OPERACION:3, ALMACEN:4, PRODUCTO:5, VARIANTE:6, CANTIDAD:7, ID_PROV:8, MARCA:9, COSTO_UNITARIO:10, MONEDA:11, IMPORTE:12, USUARIO:13, OBS:14, DNI:15, CARGO:16, FIRMA_URL:17, ESTADO:18, FECHA_CONFIRMACION:19 },
  REG:   { ID_REG:1, FECHA:2, OPERACION:3, ALMACEN:4, PRODUCTO:5, VARIANTE:6, DNI:7, NOMBRES:8, EMPRESA:9, CARGO:10, CANTIDAD:11, COSTO_UNITARIO:12, MONEDA:13, IMPORTE:14, USUARIO:15, OBS:16, DEVOLVIBLE:17, VIDA_UTIL_DIAS:18, FREC_INSP:19, REQ_CAP_TEMA:20, FECHA_VENCIMIENTO:21, PROX_INSPECCION:22, FIRMA_URL:23, REF_ID:24, ESTADO:25, FECHA_CONFIRMACION:26 },
  ALM:   { ID_ALMACEN:1, NOMBRE:2, UBICACION:3, ESTADO:4 },
  PER:   { ID:1, DNI:2, NOMBRES:3, EMPRESA:5, CARGO:7, CONDICION:12 } // PERSONAl
};

/** ================== UTILS ================== **/
function _sh(name){ return getSpreadsheetEPP().getSheetByName(name); }
function _str(x){ return (x==null)? '' : String(x).trim(); }
function _num(x){ const n=Number(x); return isNaN(n)? 0 : n; }
function _today(){ const tz=Session.getScriptTimeZone(); return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'); }
function _fmtDateOut(v){ if(!v) return ''; try{ return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }catch(e){ return _str(v); } }

function _readRows(name){
  const sh=_sh(name); if(!sh) return [];
  const vals=sh.getDataRange().getValues();
  return vals.slice(1); // sin cabecera
}
// Devuelve costo unitario en este orden: explícito → último costo (MOV) → precio en STOCK → 0
function _resolveUnitCost(almacen, producto, variante, costoUnitOpt){
  if(costoUnitOpt!=null && costoUnitOpt!=='') return _num(costoUnitOpt);
  const uc = _ultimoCosto(producto, almacen, variante);   // {unit, moneda}
  if(uc && _num(uc.unit)>0) return _num(uc.unit);
  const ps = _precioStockActual(almacen, producto, variante);
  if(_num(ps)>0) return _num(ps);
  return 0;
}

function _appendArray(name, arr){
  const sh = _sh(name);
  // fila actual antes de append
  const beforeLast = sh.getLastRow();
  sh.appendRow(arr);
  const newRow = beforeLast + 1;

  // detectar índice de columna DNI para REGISTRO y MOV (1-based)
  let dniCol = null;
  if (name === SHEPP.REGISTRO) dniCol = IDX.REG.DNI;
  else if (name === SHEPP.MOV)  dniCol = IDX.MOV.DNI;

  if (dniCol) {
    // forzamos formato texto y reescribimos el valor como string (evita pérdida de ceros)
    const cell = sh.getRange(newRow, dniCol);
    const val = String(arr[dniCol - 1] == null ? '' : arr[dniCol - 1]);
    cell.setNumberFormat('@STRING@'); // formato Texto
    cell.setValue(val);
  }
}
function _updateArrayAt(name, row1Based, arr){
  const sh=_sh(name);
  sh.getRange(row1Based, 1, 1, sh.getLastColumn()).setValues([arr]);
}
function _newRow(name){
  const sh=_sh(name); return Array(sh.getLastColumn()).fill('');
}
function _genId8(){ return Utilities.getUuid().split('-')[0]+Utilities.getUuid().split('-')[4]; }

/** ================== MATRIZ (tab + grid) ================== **/
function _readMatrizTab_(){
  const sh=_sh(SHEPP.MATRIZ); if(!sh) return [];
  const vals=sh.getDataRange().getValues();
  for (let r=0;r<Math.min(50, vals.length);r++){
    const low=vals[r].map(c=>_str(c).toLowerCase());
    if(low.includes('producto base') && low.includes('cargo')){
      const H=vals[r];
      const out=[];
      for(let i=r+1;i<vals.length;i++){
        const row=vals[i];
        if(row.every(c=>_str(c)==='')) continue;
        const obj={}; H.forEach((h,j)=> obj[h]=row[j]);
        out.push(obj);
      }
      return out;
    }
  }
  return [];
}
function _readMatrizGrid_(){
  const sh=_sh(SHEPP.MATRIZ); if(!sh) return { byProduct:{}, cargos:[], productos:[] };
  const vals=sh.getDataRange().getValues().map(r=>r.map(_str));
  const R=vals.length, C=vals[0]?.length||0;
  let prodRow = 5;
  for(let r=0;r<Math.min(15,R);r++){
    for(let c=0;c<C;c++){
      if(_str(vals[r][c]).toLowerCase()==='producto base'){ prodRow=r; break; }
    }
  }
  const capRow=Math.max(0, prodRow-4);
  const frecRow=capRow+1;
  const vidaRow=capRow+2;
  const catRow =capRow+3;
  const cargoStart = prodRow+1;

  const productos=[];
  for(let c=2;c<C;c++){
    const pb=_str(vals[prodRow][c]).trim();
    if(pb) productos.push({c, pb});
  }
  const byProduct={};
  productos.forEach(({c, pb})=>{
    const rec={
      REQ_CAP_TEMA: _str(vals[capRow]?.[c]||'').trim(),
      FREC_INSP:    _str(vals[frecRow]?.[c]||'').trim(),
      VIDA_UTIL_DIAS: _num(vals[vidaRow]?.[c]||0),
      CATEGORIA:    _str(vals[catRow]?.[c]||'').trim(),
      previstoCargos: []
    };
    byProduct[pb]=rec;
  });

  const cargos=[];
  for(let r=cargoStart;r<R;r++){
    const cargo=_str(vals[r]?.[1]||'').trim();
    if(!cargo) continue;
    cargos.push(cargo);
    productos.forEach(({c,pb})=>{
      const v=_str(vals[r]?.[c]||'').toLowerCase();
      const truthy = ['x','1','si','sí','y','yes','✔','true','ok'].includes(v);
      if(truthy) byProduct[pb].previstoCargos.push(cargo);
    });
  }
  return { byProduct, cargos, productos: productos.map(p=>p.pb) };
}

function _baseProducto(s){ return _str(s); }

function getReglasCargo(productoBase, cargo){
  const pb=_baseProducto(productoBase);
  const rows=_readMatrizTab_();
  // Busca coincidencia exacta en TAB
  for(const r of rows){
    const p=_str(r['Producto base']||r.ProductoBase||r.PRODUCTO_BASE);
    const c=_str(r.Cargo||r.CARGO);
    if(p===pb && c===_str(cargo)){
      return {
        DEVOLVIBLE: _str(r.Devolvible||r.DEVOLVIBLE||'No'),
        VIDA_UTIL_DIAS: _num(r['Vida útil (días)']||r.VidaUtilDias||r.VIDA_UTIL_DIAS||0),
        FREC_INSP: _str(r['Frec. inspección']||r.FrecInspeccion||r.FREC_INSP_DIAS||''),
        REQ_CAP_TEMA: _str(r['Req. capacitación']||r.ReqCapacitacion||r.REQ_CAPACITACION||'')
      };
    }
  }
  // Fallback a GRID por producto
  const grid=_readMatrizGrid_().byProduct[pb];
  if(grid){
    return {
      DEVOLVIBLE: 'No',
      VIDA_UTIL_DIAS: _num(grid.VIDA_UTIL_DIAS||0),
      FREC_INSP: _str(grid.FREC_INSP||''),
      REQ_CAP_TEMA: _str(grid.REQ_CAP_TEMA||'')
    };
  }
  return { DEVOLVIBLE:'No', VIDA_UTIL_DIAS:0, FREC_INSP:'', REQ_CAP_TEMA:'' };
}

function _proxFromFrec(hoyYmd, frecStr){
  if(!frecStr) return '';
  const n=_num(frecStr);
  if(n<=0) return '';
  const d=new Date(hoyYmd);
  d.setDate(d.getDate()+n);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** ================== INIT / CATÁLOGOS ================== **/
function getInit(almacenId){
  const almacenes=_readRows(SHEPP.ALMACENES).map(r=>({
    ID_ALMACEN:_str(r[IDX.ALM.ID_ALMACEN-1]),
    NOMBRE:_str(r[IDX.ALM.NOMBRE-1]),
    UBICACION:_str(r[IDX.ALM.UBICACION-1]),
    ESTADO:_str(r[IDX.ALM.ESTADO-1])
  }));

  // Derivar listas desde cache de stock (filtrado por almacén si aplica)
  const stockAll = _readStockCache_();
  const stockRows = almacenId
    ? stockAll.filter(x => _str(x.ALMACEN) === _str(almacenId))
    : stockAll;

  const productos  = [...new Set(stockRows.map(x=>x.PRODUCTO))].sort();
  const categorias = [...new Set(stockRows.map(x=>x.CATEGORIA).filter(Boolean))].sort();

  // PROVEEDOR/MARCA desde MOV (filtrado por almacén para coherencia)
  const movRowsAll=_readRows(SHEPP.MOV);
  const movRows = almacenId
    ? movRowsAll.filter(r => _str(r[IDX.MOV.ALMACEN-1]) === _str(almacenId))
    : movRowsAll;
  const proveedores=[...new Set(movRows.map(r=>_str(r[IDX.MOV.ID_PROV-1])).filter(Boolean))].sort();
  const marcas=[...new Set(movRows.map(r=>_str(r[IDX.MOV.MARCA-1])).filter(Boolean))].sort();

  // MATRIZ (sin filtro)
  const matriz=_readMatrizTab_();
  const matrizGrid=_readMatrizGrid_();

  // Nota: no enviamos "stock" completo — se usará searchStockPaged para la grilla
  return { almacenes, productos, categorias, proveedores, marcas, matriz, matrizGrid };
}


/** ============ Cache ligero del STOCK como objetos ============ **/
function _readStockCache_(){
  const cache = CacheService.getDocumentCache();
  const key = 'stock:all:v1';
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const rows = _readRows(SHEPP.STOCK);
  const list = rows.map(r=>({
    ID:_str(r[IDX.STOCK.ID-1]),
    ALMACEN:_str(r[IDX.STOCK.ALMACEN-1]),
    PRODUCTO:_str(r[IDX.STOCK.PRODUCTO-1]),
    VARIANTE:_str(r[IDX.STOCK.VARIANTE-1]),
    CATEGORIA:_str(r[IDX.STOCK.CATEGORIA-1]),
    STOCK:_num(r[IDX.STOCK.STOCK-1]),
    STOCK_MINIMO:_num(r[IDX.STOCK.STOCK_MINIMO-1]),
    PRECIO:_num(r[IDX.STOCK.PRECIO-1]),
    IMG:_str(r[IDX.STOCK.IMG-1]),
    FILE:_str(r[IDX.STOCK.FILE-1])
  }));
  // TTL 120s
  cache.put(key, JSON.stringify(list), 120);
  return list;
}


/** ============ Paginación Server-Side (30 en 30) ============ **/
function searchStockPaged(payload){
  payload = payload || {};
  const almacenId = _str(payload.almacenId||'');
  const term      = _str(payload.term||'').toLowerCase();
  const cat       = _str(payload.cat||'Todos');
  const onlyStock = !!payload.onlyStock;
  const offset    = Math.max(0, Number(payload.offset||0));
  const limitReq  = Math.max(0, Number(payload.limit||30)); // 0 => sin límite (export)

  // 1) base: cache de STOCK
  let list = _readStockCache_();

  // 2) filtros
  if (almacenId) list = list.filter(x=> _str(x.ALMACEN)===almacenId);
  if (cat && cat!=='Todos') list = list.filter(x=> _str(x.CATEGORIA)===cat);
  if (term){
    const t = term;
    list = list.filter(x=>{
      return (x.PRODUCTO||'').toLowerCase().includes(t) ||
             (x.VARIANTE||'').toLowerCase().includes(t) ||
             (x.CATEGORIA||'').toLowerCase().includes(t);
    });
  }
  if (onlyStock) list = list.filter(x=> Number(x.STOCK||0) > 0);

  // 3) ordenar por PRODUCTO y luego VARIANTE
  list.sort((a,b)=>{
    const pa = _str(a.PRODUCTO), pb = _str(b.PRODUCTO);
    if (pa!==pb) return pa.localeCompare(pb);
    const va = _str(a.VARIANTE), vb = _str(b.VARIANTE);
    return va.localeCompare(vb);
  });

  const total = list.length;

  // 4) slicing "product-aware": no cortar un producto en el borde de página
  if (limitReq===0){
    // exportar todo lo filtrado
    return { total, rows: list.slice(offset) };
  }

  const start = Math.min(offset, total);
  let end = Math.min(total, offset + limitReq);
  let page = list.slice(start, end);

  if (page.length > 0 && end < total){
    const lastProd = _str(page[page.length - 1].PRODUCTO);
    // extender mientras el siguiente registro siga siendo el mismo producto
    while (end < total && _str(list[end].PRODUCTO) === lastProd){
      page.push(list[end]);
      end++;
    }
  }

  return { total, rows: page };
}


/** ============ Variantes por producto (para modal Movimiento) ============ **/
function getVariantesPorProducto(almacenId, producto){
  const list = _readStockCache_();
  const rows = list.filter(x=> _str(x.ALMACEN)===_str(almacenId) && _str(x.PRODUCTO)===_str(producto));
  const vars = [...new Set(rows.map(x=>_str(x.VARIANTE||'')).filter(Boolean))].sort();
  return vars;
}


/** ================= INVALIDAR CACHÉS (stock + historial) ================= */
function _invalidateStockCache_(){
  try{
    const cache = CacheService.getDocumentCache();
    ['stock:all:v1','registro:all:v1','mov:all:v1'].forEach(k=>{
      try{ cache.remove(k); }catch(_){}
    });
  }catch(e){
    console.warn('No se pudo invalidar cachés:', e);
  }
}
/** ============ Cache ligero del REGISTRO como objetos ============ **/
function _readRegistroCache_(){
  const cache = CacheService.getDocumentCache();
  const key = 'registro:all:v1';
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const rows = _readRows(SHEPP.REGISTRO); // sin cabecera
  const list = rows.map(r=>({
    ID       : _str(r[IDX.REG.ID_REG-1]),
    FECHA    : _str(r[IDX.REG.FECHA-1]),
    OPERACION: _str(r[IDX.REG.OPERACION-1]),
    ALMACEN  : _str(r[IDX.REG.ALMACEN-1]),
    PRODUCTO : _str(r[IDX.REG.PRODUCTO-1]),
    VARIANTE : _str(r[IDX.REG.VARIANTE-1]),
    DNI      : _str(r[IDX.REG.DNI-1]),
    NOMBRES  : _str(r[IDX.REG.NOMBRES-1]),
    EMPRESA  : _str(r[IDX.REG.EMPRESA-1]),
    CARGO    : _str(r[IDX.REG.CARGO-1]),
    CANTIDAD : _num(r[IDX.REG.CANTIDAD-1]),
    COSTO_UNITARIO: _num(r[IDX.REG.COSTO_UNITARIO-1]),
    MONEDA   : _str(r[IDX.REG.MONEDA-1]),
    IMPORTE  : _num(r[IDX.REG.IMPORTE-1]),
    USUARIO  : _str(r[IDX.REG.USUARIO-1]),
    OBS      : _str(r[IDX.REG.OBS-1]),
    FIRMA_URL: _str(r[IDX.REG.FIRMA_URL-1]),
    REF_ID   : _str(r[IDX.REG.REF_ID-1]),
    ESTADO   : _str(r[IDX.REG.ESTADO-1] || ''),
    FECHA_CONFIRMACION: _str(r[IDX.REG.FECHA_CONFIRMACION-1] || '')
  }));
  // TTL 120s
  cache.put(key, JSON.stringify(list), 120);
  return list;
}
/** ============ Cache ligero de MOV como objetos ============ **/
function _readMovCache_(){
  const cache = CacheService.getDocumentCache();
  const key = 'mov:all:v1';
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const rows = _readRows(SHEPP.MOV);
  const list = rows.map(r=>({
    ID       : _str(r[IDX.MOV.ID_MOV-1]),
    FECHA    : _str(r[IDX.MOV.FECHA-1]),
    OPERACION: _str(r[IDX.MOV.OPERACION-1]),
    ALMACEN  : _str(r[IDX.MOV.ALMACEN-1]),
    PRODUCTO : _str(r[IDX.MOV.PRODUCTO-1]),
    VARIANTE : _str(r[IDX.MOV.VARIANTE-1]),
    CANTIDAD : _num(r[IDX.MOV.CANTIDAD-1]),
    ID_PROV  : _str(r[IDX.MOV.ID_PROV-1]),
    MARCA    : _str(r[IDX.MOV.MARCA-1]),
    COSTO_UNITARIO: _num(r[IDX.MOV.COSTO_UNITARIO-1]),
    MONEDA   : _str(r[IDX.MOV.MONEDA-1]),
    IMPORTE  : _num(r[IDX.MOV.IMPORTE-1]),
    USUARIO  : _str(r[IDX.MOV.USUARIO-1]),
    OBS      : _str(r[IDX.MOV.OBS-1]),
    DNI      : typeof IDX.MOV.DNI==='number' ? _str(r[IDX.MOV.DNI-1]) : '',
    CARGO    : typeof IDX.MOV.CARGO==='number' ? _str(r[IDX.MOV.CARGO-1]) : '',
    FIRMA_URL: typeof IDX.MOV.FIRMA_URL==='number' ? _str(r[IDX.MOV.FIRMA_URL-1]) : '',
    ESTADO   : typeof IDX.MOV.ESTADO==='number' ? _str(r[IDX.MOV.ESTADO-1]) : '',
    FECHA_CONFIRMACION: typeof IDX.MOV.FECHA_CONFIRMACION==='number' ? _str(r[IDX.MOV.FECHA_CONFIRMACION-1]) : ''
  }));
  // TTL 120s
  cache.put(key, JSON.stringify(list), 120);
  return list;
}
/** ============ Historial paginado (SSR 30x30) ============ **/
function searchHistorialPaged(payload){
  payload = payload || {};
  const tabla     = _str(payload.tabla||'REGISTRO').toUpperCase(); // 'REGISTRO' | 'MOV'
  const almacenId = _str(payload.almacenId||'');  // filtro exacto
  const dni       = _str(payload.dni||'');        // filtro exacto
  const term      = _str(payload.term||'').toLowerCase(); // texto libre
  const ops       = Array.isArray(payload.ops) ? payload.ops.map(s=>_str(s).toLowerCase()) : [];
  const dateFrom  = _str(payload.date_from||'');  // 'YYYY-MM-DD' (opcional)
  const dateTo    = _str(payload.date_to||'');    // 'YYYY-MM-DD' (opcional)
  const offset    = Math.max(0, Number(payload.offset||0));
  const limit     = Math.max(0, Number(payload.limit||30)); // 0 => sin límite

  const tz = Session.getScriptTimeZone();
  const toMs = (v)=>{
    if (v instanceof Date) return v.getTime();
    const s=_str(v); if(!s) return 0;
    // si viene 'YYYY-MM-DD'
    const parts = s.split('-');
    if (parts.length===3){
      const d = new Date(Number(parts[0]), Number(parts[1])-1, Number(parts[2]));
      return d.getTime();
    }
    return new Date(s).getTime();
  };

  let list = (tabla==='MOV') ? _readMovCache_() : _readRegistroCache_();

  // === Filtros ===
  if (almacenId) list = list.filter(x => _str(x.ALMACEN) === almacenId);
  if (dni)       list = list.filter(x => _str(x.DNI) === dni);

  if (ops.length){
    const setOps = new Set(ops);
    list = list.filter(x => setOps.has(_str(x.OPERACION).toLowerCase()));
  }

  if (term){
    const t = term;
    list = list.filter(x=>{
      return (_str(x.PRODUCTO).toLowerCase().includes(t)) ||
             (_str(x.VARIANTE).toLowerCase().includes(t)) ||
             (_str(x.OBS).toLowerCase().includes(t)) ||
             (_str(x.EMPRESA).toLowerCase().includes(t)) ||
             (_str(x.CARGO).toLowerCase().includes(t)) ||
             (_str(x.NOMBRES).toLowerCase().includes(t));
    });
  }

  let fromMs = dateFrom ? toMs(dateFrom) : null;
  let toMsInc= dateTo   ? (toMs(dateTo) + 86400000 - 1) : null; // inclusivo hasta el final del día
  if (fromMs!=null || toMsInc!=null){
    list = list.filter(x=>{
      const ms = toMs(x.FECHA);
      if (fromMs!=null && ms < fromMs) return false;
      if (toMsInc!=null && ms > toMsInc) return false;
      return true;
    });
  }

  // Orden por fecha DESC (reciente primero), usando FECHA y luego ID
  list.sort((a,b)=>{
    const da = toMs(a.FECHA), db = toMs(b.FECHA);
    if (db!==da) return db-da;
    return _str(b.ID).localeCompare(_str(a.ID)); // estabilidad
  });

  const total = list.length;
  const page  = limit>0 ? list.slice(offset, offset+limit) : list;

  return { total, rows: page };
}

/** ================== PERSONA ================== **/
function getPersonaByDni(dni){
  const sh = getSpreadsheetPersonal().getSheetByName('PERSONAL');
  if (!sh) return null;

  const vals = sh.getDataRange().getValues(); // incluye cabecera en vals[0]
  for (let i = 1; i < vals.length; i++) {     // empezamos en 1 para omitir cabecera
    const r = vals[i];
    if (_str(r[IDX.PER.DNI-1]) === _str(dni)) {
      return {
        ID:        _str(r[IDX.PER.ID-1]||''),
        DNI:       _str(r[IDX.PER.DNI-1]||''),
        NOMBRES:   _str(r[IDX.PER.NOMBRES-1]||''),
        EMPRESA:   _str(r[IDX.PER.EMPRESA-1]||''),
        CARGO:     _str(r[IDX.PER.CARGO-1]||''),
        CONDICION: _str(r[IDX.PER.CONDICION-1]||'')
      };
    }
  }
  return null;
}





/**
 * Busca persona por DNI, nombre o apellido (flexible)
 * @param {string} query - Texto de búsqueda
 * @returns {Object} { ok, persona, multiple, personas, message }
 */
function buscarPersonaFlexible(query) {
  try {
    const sh = getSpreadsheetPersonal().getSheetByName('PERSONAL');
    if (!sh) return { ok: false, message: 'Hoja PERSONAL no encontrada' };

    const vals = sh.getDataRange().getValues();
    const q = _str(query).toLowerCase();
    const matches = [];

    for (let i = 1; i < vals.length; i++) {
      const r = vals[i];
      const dni = _str(r[IDX.PER.DNI-1]);
      const nombres = _str(r[IDX.PER.NOMBRES-1]).toLowerCase();
      
      // Coincidencia por DNI exacto o nombre parcial
      if (dni === query || nombres.includes(q)) {
        matches.push({
          ID:        _str(r[IDX.PER.ID-1]||''),
          DNI:       dni,
          NOMBRES:   _str(r[IDX.PER.NOMBRES-1]||''),
          EMPRESA:   _str(r[IDX.PER.EMPRESA-1]||''),
          CARGO:     _str(r[IDX.PER.CARGO-1]||''),
          CONDICION: _str(r[IDX.PER.CONDICION-1]||'')
        });
      }
    }

    // 1 resultado: devolver directamente
    if (matches.length === 1) {
      return { ok: true, persona: matches[0] };
    }

    // Múltiples: devolver lista
    if (matches.length > 1) {
      return { ok: true, multiple: true, personas: matches };
    }

    // Sin resultados
    return { ok: false, message: 'Sin resultados' };

  } catch (e) {
    return { ok: false, message: e.message || String(e) };
  }
}

/** ================== STOCK HELPERS ================== **/
function _precioStockActual(almacen, producto, variante){
  const rows=_readRows(SHEPP.STOCK);
  const rec=rows.find(r=>
    _str(r[IDX.STOCK.ALMACEN-1])===_str(almacen) &&
    _str(r[IDX.STOCK.PRODUCTO-1])===_str(producto) &&
    _str(r[IDX.STOCK.VARIANTE-1]||'')===_str(variante||'')
  );
  return rec? _num(rec[IDX.STOCK.PRECIO-1]||0) : 0;
}

// Aumenta/disminuye stock; crea si falta y createIfMissing=true. Aplica "patch" en columnas (precio, min, img, file, categoria).
function _ensureStockDelta(almacen, producto, variante, delta, createIfMissing, patch){
  const sh=_sh(SHEPP.STOCK);
  const all=sh.getDataRange().getValues(); // con cabecera
  const H=all[0], rows=all.slice(1);
  let idxRow=-1;
  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    if(_str(r[IDX.STOCK.ALMACEN-1])===_str(almacen) &&
       _str(r[IDX.STOCK.PRODUCTO-1])===_str(producto) &&
       _str(r[IDX.STOCK.VARIANTE-1]||'')===_str(variante||'')){
      idxRow = i+2; // +2 por cabecera + 1-based
      break;
    }
  }
  if(idxRow<0){
    if(!createIfMissing) throw new Error('No existe en STOCK esa combinación (ALMACEN, PRODUCTO, VARIANTE)');
    const row=_newRow(SHEPP.STOCK);
    row[IDX.STOCK.ID-1]       = Utilities.getUuid();
    row[IDX.STOCK.ALMACEN-1]  = almacen;
    row[IDX.STOCK.PRODUCTO-1] = producto;
    row[IDX.STOCK.VARIANTE-1] = variante||'';
    row[IDX.STOCK.CATEGORIA-1]= (patch && patch.CATEGORIA)? patch.CATEGORIA : '';
    row[IDX.STOCK.IMG-1]      = (patch && patch.IMG) || '';
    row[IDX.STOCK.FILE-1]     = (patch && patch.FILE) || '';
    row[IDX.STOCK.STOCK-1]    = Math.max(0,_num(delta||0));
    row[IDX.STOCK.STOCK_MINIMO-1] = (patch && patch.hasOwnProperty('STOCK_MINIMO')) ? _num(patch.STOCK_MINIMO) : 0;
    row[IDX.STOCK.PRECIO-1]   = (patch && patch.hasOwnProperty('PRECIO')) ? _num(patch.PRECIO) : 0;
    _appendArray(SHEPP.STOCK, row);
    return;
  }
  const row = sh.getRange(idxRow, 1, 1, H.length).getValues()[0];
  const nuevo= _num(row[IDX.STOCK.STOCK-1]) + _num(delta||0);
  if(nuevo<0) throw new Error('Stock insuficiente');
  row[IDX.STOCK.STOCK-1]=nuevo;
  if(patch && typeof patch==='object'){
    if(patch.hasOwnProperty('PRECIO'))       row[IDX.STOCK.PRECIO-1] = _num(patch.PRECIO);
    if(patch.hasOwnProperty('STOCK_MINIMO')) row[IDX.STOCK.STOCK_MINIMO-1] = _num(patch.STOCK_MINIMO);
    if(patch.hasOwnProperty('IMG'))          row[IDX.STOCK.IMG-1] = _str(patch.IMG);
    if(patch.hasOwnProperty('FILE'))         row[IDX.STOCK.FILE-1] = _str(patch.FILE);
    if(patch.hasOwnProperty('CATEGORIA'))    row[IDX.STOCK.CATEGORIA-1] = _str(patch.CATEGORIA);
  }
  _updateArrayAt(SHEPP.STOCK, idxRow, row);
}

/** ================== COSTO / ÚLTIMO COSTO ================== **/
function _ultimoCosto(producto, almacen, variante){
  const rows=_readRows(SHEPP.MOV)
    .filter(r=> _str(r[IDX.MOV.PRODUCTO-1])===_str(producto))
    .sort((a,b)=> new Date(_str(b[IDX.MOV.FECHA-1])) - new Date(_str(a[IDX.MOV.FECHA-1])));
  const byExact= rows.find(x=>
    _str(x[IDX.MOV.ALMACEN-1])===_str(almacen) &&
    _str(x[IDX.MOV.VARIANTE-1]||'')===_str(variante||'') &&
    _str(x[IDX.MOV.COSTO_UNITARIO-1])!==''
  );
  if(byExact) return { unit:_num(byExact[IDX.MOV.COSTO_UNITARIO-1]), moneda:_str(byExact[IDX.MOV.MONEDA-1]||'PEN') };
  const byProd= rows.find(x=> _str(x[IDX.MOV.ALMACEN-1])===_str(almacen) && _str(x[IDX.MOV.COSTO_UNITARIO-1])!=='' );
  if(byProd) return { unit:_num(byProd[IDX.MOV.COSTO_UNITARIO-1]), moneda:_str(byProd[IDX.MOV.MONEDA-1]||'PEN') };
  const anyVar= rows.find(x=> _str(x[IDX.MOV.COSTO_UNITARIO-1])!=='' );
  if(anyVar) return { unit:_num(anyVar[IDX.MOV.COSTO_UNITARIO-1]), moneda:_str(anyVar[IDX.MOV.MONEDA-1]||'PEN') };
  return { unit:0, moneda:'PEN' };
}
function getUltimoCosto(producto, almacen, variante){ return _ultimoCosto(producto, almacen, variante); }

/** ================== REGISTRO (Entrega/Devolución) + espejo en MOV ================== **/
function registrarEntrega(payload){
  const lock=LockService.getScriptLock(); lock.tryLock(30000);
  try{
    const cant=_num(payload.cantidad||0); 
    if(cant<=0) throw new Error('Cantidad debe ser > 0');
    
    const producto=_str(payload.producto), almacen=_str(payload.almacen);
    const variante=_str(payload.variante||'');
    const base=producto;

    const reglas=getReglasCargo(base, _str(payload.cargo));
    if(payload.devolvible!=null && payload.devolvible!==''){
      reglas.DEVOLVIBLE = String(payload.devolvible).toLowerCase().startsWith('s') || payload.devolvible===true ? 'Sí':'No';
    }

    let precioUnit = (payload.precioUnit!=null && payload.precioUnit!=='') ? _num(payload.precioUnit) : null;
    let moneda = _str(payload.moneda||'');
    if(precioUnit==null){ 
      const uc=_ultimoCosto(producto, almacen, variante); 
      precioUnit=uc.unit; 
      if(!moneda) moneda=uc.moneda; 
    }
    if(!moneda) moneda='PEN';

    const hoy = _today();
    
    let vence = '';
    if (reglas.VIDA_UTIL_DIAS > 0) {
      const fVence = new Date(hoy);
      fVence.setDate(fVence.getDate() + reglas.VIDA_UTIL_DIAS);
      vence = Utilities.formatDate(fVence, 'GMT-5', 'yyyy-MM-dd');
    }
    
    const prox = _proxFromFrec(hoy, reglas.FREC_INSP);

    // 🆕 DETECTAR RETRASO (antes de crear el row)
    const ultimaEntrega = buscarUltimaEntrega(payload.dni, producto, variante);
    let obsAuto = _str(payload.obs || '');
    
    if (ultimaEntrega && ultimaEntrega.FECHA_VENCIMIENTO) {
      const vencPrev = new Date(ultimaEntrega.FECHA_VENCIMIENTO);
      const hoyDate = new Date(hoy);
      
      if (hoyDate > vencPrev) {
        const diffMs = hoyDate - vencPrev;
        const diasRetraso = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
        const retrasoMsg = `⚠️ Reemplazo con retraso de ${diasRetraso} día${diasRetraso===1?'':'s'}.`;
        obsAuto = obsAuto ? `${retrasoMsg} ${obsAuto}` : retrasoMsg;
      } else {
        const aTiempoMsg = '✅ Reemplazo a tiempo.';
        obsAuto = obsAuto ? `${aTiempoMsg} ${obsAuto}` : aTiempoMsg;
      }
    }

    // 🔹 REGISTRO (hoja REGISTRO)
    const row=_newRow(SHEPP.REGISTRO);
    row[IDX.REG.ID_REG-1]=_genId8();
    row[IDX.REG.FECHA-1]=hoy;
    row[IDX.REG.OPERACION-1]='Entrega';
    row[IDX.REG.ALMACEN-1]=almacen;
    row[IDX.REG.PRODUCTO-1]=producto;
    row[IDX.REG.VARIANTE-1]=variante;
    row[IDX.REG.DNI-1]=_str(payload.dni||'');
    row[IDX.REG.NOMBRES-1]=_str(payload.nombres||'');
    row[IDX.REG.EMPRESA-1]=_str(payload.empresa||'');
    row[IDX.REG.CARGO-1]=_str(payload.cargo||'');
    row[IDX.REG.CANTIDAD-1]=cant;
    row[IDX.REG.COSTO_UNITARIO-1]=precioUnit;
    row[IDX.REG.MONEDA-1]=moneda;
    row[IDX.REG.IMPORTE-1]=precioUnit*cant;
    row[IDX.REG.USUARIO-1]=Session.getActiveUser().getEmail();
    row[IDX.REG.OBS-1]=obsAuto; // ← 🆕 Observación con detección de retraso
    row[IDX.REG.DEVOLVIBLE-1]=reglas.DEVOLVIBLE;
    row[IDX.REG.VIDA_UTIL_DIAS-1]=reglas.VIDA_UTIL_DIAS;
    row[IDX.REG.FREC_INSP-1]=reglas.FREC_INSP;
    row[IDX.REG.REQ_CAP_TEMA-1]=reglas.REQ_CAP_TEMA;
    row[IDX.REG.FECHA_VENCIMIENTO-1]=vence||'';
    row[IDX.REG.PROX_INSPECCION-1]=prox||'';
    row[IDX.REG.FIRMA_URL-1]=''; // Firma se llena cuando el trabajador confirme
    row[IDX.REG.REF_ID-1]=_str(payload.refId||'');
    row[IDX.REG.ESTADO-1]='Pendiente';
    row[IDX.REG.FECHA_CONFIRMACION-1]='';
    _appendArray(SHEPP.REGISTRO, row);

    // 🔹 MOVIMIENTOS (hoja MOVIMIENTOS)
    const m=_newRow(SHEPP.MOV);
    m[IDX.MOV.ID_MOV-1]=_genId8();
    m[IDX.MOV.FECHA-1]=hoy;
    m[IDX.MOV.OPERACION-1]='Entrega';
    m[IDX.MOV.ALMACEN-1]=almacen;
    m[IDX.MOV.PRODUCTO-1]=producto;
    m[IDX.MOV.VARIANTE-1]=variante;
    m[IDX.MOV.CANTIDAD-1]=cant;
    m[IDX.MOV.ID_PROV-1]='';
    m[IDX.MOV.MARCA-1]='';
    m[IDX.MOV.COSTO_UNITARIO-1]=precioUnit;
    m[IDX.MOV.MONEDA-1]=moneda;
    m[IDX.MOV.IMPORTE-1]=precioUnit*cant;
    m[IDX.MOV.USUARIO-1]=Session.getActiveUser().getEmail();
    m[IDX.MOV.OBS-1]=obsAuto;
    m[IDX.MOV.DNI-1]=_str(payload.dni||'');
    m[IDX.MOV.CARGO-1]=_str(payload.cargo||'');
    m[IDX.MOV.FIRMA_URL-1]='';
    m[IDX.MOV.ESTADO-1]='Pendiente';
    m[IDX.MOV.FECHA_CONFIRMACION-1]='';
    _appendArray(SHEPP.MOV, m);

    // 🚫 YA NO se descuenta stock aquí — se descuenta cuando el trabajador confirme

    return { ok:true, id: row[IDX.REG.ID_REG-1] };
  }catch(err){
    return { ok:false, message: err.message||String(err) };
  }finally{
    try{ invalidateStockCache(); }catch(e){}
    try{ lock.releaseLock(); }catch(e){}
  }
}



function execCart(payload){
  try{
    const almacen = _str(payload.almacen||'');
    if(!almacen) throw new Error('Almacén requerido');

    const dni      = _str(payload.dni||'');
    const nombres  = _str(payload.nombres||'');
    const empresa  = _str(payload.empresa||'');
    const cargo    = _str(payload.cargo||'');
    const firmaUrl = _str(payload.firmaUrl||'');

    const items = Array.isArray(payload.items) ? payload.items : [];
    if(!items.length) return { ok:false, message:'Carrito vacío' };

    const doneIds = [];
    const errors  = [];

    for (const it of items){
      const tipo = _str(it.tipo||'').toLowerCase();
      const base = {
        producto   : _str(it.producto),
        variante   : _str(it.variante||''),
        cantidad   : _num(it.cantidad||0),
        precioUnit : (it.precioUnit!=null && it.precioUnit!=='') ? _num(it.precioUnit) : null,
        moneda     : _str(it.moneda||'PEN'),
        obs        : _str(it.obs||''),
        almacen,
        dni, nombres, empresa, cargo,
        firmaUrl,
        // 👇 Nuevos campos enviados desde frontend
        vida_util  : _num(it.vida_util || 0),
        freq       : _str(it.freq || ''),
        devolvible : (String(it.devolvible||'').toLowerCase().startsWith('s') || it.devolvible === true) ? 'Sí' : 'No'
      };

      if(!base.producto) { errors.push('Producto requerido'); continue; }
      if(base.cantidad<=0){ errors.push('Cantidad debe ser > 0'); continue; }

      let r;
      try{
        if (tipo.startsWith('entreg')) {
          r = registrarEntrega(base);
        } else if (tipo.startsWith('devol')) {
          r = registrarDevolucion(base);
        } else {
          throw new Error('Tipo no soportado en carrito: ' + it.tipo);
        }

        if (r && r.ok){
          doneIds.push(r.id);
        } else {
          errors.push(r?.message || 'Error al registrar');
        }
      }catch(e){
        errors.push(String(e && e.message ? e.message : e));
      }
    }

    if (errors.length){
      return { ok:false, message:'Algunas operaciones fallaron', errors, doneIds };
    }

    return { ok:true, ids: doneIds };

  }catch(err){
    return { ok:false, message: err.message || String(err) };
  }finally{
    try{ invalidateStockCache(); }catch(_){}
  }
}


function registrarDevolucion(payload){
  const lock=LockService.getScriptLock(); lock.tryLock(30000);
  try{
    const cant=_num(payload.cantidad||0); if(cant<=0) throw new Error('Cantidad debe ser > 0');
    const producto=_str(payload.producto), almacen=_str(payload.almacen);
    const variante=_str(payload.variante||'');

    // 💡 NUEVO: resolver categoría para la posible creación de fila en STOCK
    const categoria = _resolveCategoria(almacen, producto, variante, _str(payload.categoria||''));

    let precioUnit=(payload.precioUnit!=null && payload.precioUnit!=='')? _num(payload.precioUnit): null;
    let moneda=_str(payload.moneda||'');
    if(precioUnit==null){ const uc=_ultimoCosto(producto, almacen, variante); precioUnit=uc.unit; if(!moneda) moneda=uc.moneda; }
    if(!moneda) moneda='PEN';

    const hoy=_today();
    // REGISTRO
    const row=_newRow(SHEPP.REGISTRO);
    row[IDX.REG.ID_REG-1]=_genId8();
    row[IDX.REG.FECHA-1]=hoy;
    row[IDX.REG.OPERACION-1]='Devolución';
    row[IDX.REG.ALMACEN-1]=almacen;
    row[IDX.REG.PRODUCTO-1]=producto;
    row[IDX.REG.VARIANTE-1]=variante;
    row[IDX.REG.DNI-1]=_str(payload.dni||'');
    row[IDX.REG.NOMBRES-1]=_str(payload.nombres||'');
    row[IDX.REG.EMPRESA-1]=_str(payload.empresa||'');
    row[IDX.REG.CARGO-1]=_str(payload.cargo||'');
    row[IDX.REG.CANTIDAD-1]=cant;
    row[IDX.REG.COSTO_UNITARIO-1]=precioUnit;
    row[IDX.REG.MONEDA-1]=moneda;
    row[IDX.REG.IMPORTE-1]=precioUnit*cant;
    row[IDX.REG.USUARIO-1]=Session.getActiveUser().getEmail();
    row[IDX.REG.OBS-1]=_str(payload.obs||'');
    row[IDX.REG.FIRMA_URL-1]=_str(payload.firmaUrl||'');
    row[IDX.REG.REF_ID-1]=_str(payload.refId||'');
    _appendArray(SHEPP.REGISTRO, row);

    // MOV
    const m=_newRow(SHEPP.MOV);
    m[IDX.MOV.ID_MOV-1]=_genId8();
    m[IDX.MOV.FECHA-1]=hoy;
    m[IDX.MOV.OPERACION-1]='Devolución';
    m[IDX.MOV.ALMACEN-1]=almacen;
    m[IDX.MOV.PRODUCTO-1]=producto;
    m[IDX.MOV.VARIANTE-1]=variante;
    m[IDX.MOV.CANTIDAD-1]=cant;
    m[IDX.MOV.ID_PROV-1]=''; m[IDX.MOV.MARCA-1]='';
    m[IDX.MOV.COSTO_UNITARIO-1]=precioUnit; m[IDX.MOV.MONEDA-1]=moneda; m[IDX.MOV.IMPORTE-1]=precioUnit*cant;
    m[IDX.MOV.USUARIO-1]=Session.getActiveUser().getEmail();
    m[IDX.MOV.OBS-1]=_str(payload.obs||'');
    m[IDX.MOV.DNI-1]=_str(payload.dni||'');
    m[IDX.MOV.CARGO-1]=_str(payload.cargo||'');
    m[IDX.MOV.FIRMA_URL-1]=_str(payload.firmaUrl||'');
    _appendArray(SHEPP.MOV, m);

    // 👉 Si la variante no existía, la creamos con su categoría
    _ensureStockDelta(almacen, producto, variante, +cant, true, { CATEGORIA: categoria });

    return { ok:true, id: row[IDX.REG.ID_REG-1] };
  }catch(err){
    return { ok:false, message: err.message||String(err) };
  }finally{
    try{ invalidateStockCache(); }catch(e){}
    try{ lock.releaseLock(); }catch(e){}
  }
}

/** ================== Movimiento (Ingreso/Transferencia/Ajuste/Baja) ================== **/
function registrarMovimiento(payload){
  const lock=LockService.getScriptLock(); lock.tryLock(30000);
  try{
    // ====== INPUTS ======
    const rawTipo = _str(payload.tipo||'').trim();
    const producto=_str(payload.producto||''); if(!producto) throw new Error('Producto requerido');
    const almacen=_str(payload.almacen||'');  if(!almacen)  throw new Error('Almacén requerido');
    const variante=_str(payload.variante||'');
    const cant=_num(payload.cantidad||0);
    const proveedor=_str(payload.proveedor||'');
    const marca=_str(payload.marca||'');
    const costoUnitOpt = (payload.costoUnit===undefined || payload.costoUnit==='') ? null : _num(payload.costoUnit);
    const moneda=_str(payload.moneda||'PEN');
    const obs=_str(payload.obs||'');
    const hoy=_today();
    const usuario=Session.getActiveUser().getEmail();

    // 💡 NUEVO: resolver categoría (opcionalmente payload.categoria puede forzarla)
    const categoria = _resolveCategoria(almacen, producto, variante, _str(payload.categoria||''));

    // ====== NORMALIZAR TIPO ======
    function norm(s){
      const t=_str(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,''); // sin acentos
      const c=t.replace(/\s+/g,'');
      if(/^ingreso|^entrada/.test(c)) return 'INGRESO';
      if(/^baja|^salida/.test(c))     return 'BAJA';
      if(/^ajuste\+$|^ajusteplus$/.test(c)) return 'AJUSTE_PLUS';
      if(/^ajuste\-$/ .test(c) || /^ajusteminus$/.test(c)) return 'AJUSTE_MINUS';
      if(/^ajuste/.test(c))          return 'AJUSTE';
      if(/^transfer/.test(c) || /^transf/.test(c)) return 'TRANSFERENCIA';
      return 'OTRO';
    }
    const tipo = norm(rawTipo);

    function _getStockRec(alm, prod, var_){
      const sh=_sh(SHEPP.STOCK);
      const all=sh.getDataRange().getValues(); // con cabecera
      for (let i=1;i<all.length;i++){
        const r=all[i];
        if(_str(r[IDX.STOCK.ALMACEN-1])===_str(alm) &&
           _str(r[IDX.STOCK.PRODUCTO-1])===_str(prod) &&
           _str(r[IDX.STOCK.VARIANTE-1]||'')===_str(var_||'')){
          return { row:r, idx:i+1 };
        }
      }
      return null;
    }
    function _appendMov(op, alm, deltaAbs, pUnit, mon){
      const m=_newRow(SHEPP.MOV);
      m[IDX.MOV.ID_MOV-1]=_genId8();
      m[IDX.MOV.FECHA-1]=hoy;
      m[IDX.MOV.OPERACION-1]=op;
      m[IDX.MOV.ALMACEN-1]=alm;
      m[IDX.MOV.PRODUCTO-1]=producto;
      m[IDX.MOV.VARIANTE-1]=variante;
      m[IDX.MOV.CANTIDAD-1]=Math.abs(deltaAbs);
      m[IDX.MOV.ID_PROV-1]=proveedor;
      m[IDX.MOV.MARCA-1]=marca;
      m[IDX.MOV.COSTO_UNITARIO-1]=_num(pUnit||0);
      m[IDX.MOV.MONEDA-1]=mon||'PEN';
      m[IDX.MOV.IMPORTE-1]=_num(pUnit||0)*Math.abs(deltaAbs);
      m[IDX.MOV.USUARIO-1]=usuario;
      m[IDX.MOV.OBS-1]=obs;
      _appendArray(SHEPP.MOV, m);
    }
    const unitAt = (alm)=> _resolveUnitCost(alm, producto, variante, costoUnitOpt);

    // ====== CASOS ======
    if(tipo==='INGRESO'){
      if(cant<=0) throw new Error('Cantidad debe ser > 0');
      const unit = unitAt(almacen);
      _appendMov('Ingreso', almacen, +cant, unit, moneda);
      // 👉 ahora fijamos CATEGORIA cuando se crea una variante nueva
      _ensureStockDelta(almacen, producto, variante, +cant, true, {PRECIO:unit, CATEGORIA: categoria});
      return { ok:true };
    }

    if(tipo==='BAJA'){
      if(cant<=0) throw new Error('Cantidad debe ser > 0');
      const unit = unitAt(almacen);
      _appendMov('Baja', almacen, -cant, unit, moneda);
      _ensureStockDelta(almacen, producto, variante, -cant, false, null);
      return { ok:true };
    }

    if(tipo==='AJUSTE'){
      // Ajuste a valor absoluto: requiere nuevoStock
      const nuevoStock=_num(payload.nuevoStock);
      if(isNaN(nuevoStock)) throw new Error('nuevoStock requerido para Ajuste');
      const ori=_getStockRec(almacen, producto, variante);
      const actual = ori? _num(ori.row[IDX.STOCK.STOCK-1]) : 0;
      const delta = nuevoStock - actual;
      if(delta===0) return { ok:true, message:'Sin cambios' };
      const unit = unitAt(almacen);
      _appendMov('Ajuste', almacen, delta, unit, moneda);
      // 👉 incluir CATEGORIA si la fila se crea
      _ensureStockDelta(almacen, producto, variante, delta, true, {PRECIO:unit, CATEGORIA: categoria});
      return { ok:true };
    }

    if(tipo==='AJUSTE_PLUS' || tipo==='AJUSTE_MINUS'){
      // Ajuste por delta usando "cantidad" (+/-)
      if(cant<=0) throw new Error('Cantidad debe ser > 0');
      const delta = (tipo==='AJUSTE_PLUS') ? +cant : -cant;
      const unit = unitAt(almacen);
      _appendMov(tipo==='AJUSTE_PLUS' ? 'Ajuste +' : 'Ajuste -', almacen, delta, unit, moneda);
      // 👉 incluir CATEGORIA si la fila se crea
      _ensureStockDelta(almacen, producto, variante, delta, true, {PRECIO:unit, CATEGORIA: categoria});
      return { ok:true };
    }

    if(tipo==='TRANSFERENCIA'){
      const destino = _str(
        payload.destino ?? payload.destinoAlmacen ?? payload.almacenDestino ??
        payload.destino_id ?? payload.almacenDest ?? payload.almacen_dest ??
        payload.almDest ?? payload.to ?? ''
      );
      if(!destino) throw new Error('Destino requerido');
      if(cant<=0) throw new Error('Cantidad debe ser > 0');

      const ori=_getStockRec(almacen, producto, variante);
      const stockOri = ori? _num(ori.row[IDX.STOCK.STOCK-1]) : 0;
      if(stockOri < cant) throw new Error('Stock insuficiente en origen');

      const unitDest = unitAt(destino);
      const metaPatch = {
        PRECIO:       unitDest,
        STOCK_MINIMO: ori? _num(ori.row[IDX.STOCK.STOCK_MINIMO-1]||0) : 0,
        IMG:          ori? _str(ori.row[IDX.STOCK.IMG-1]||'') : '',
        FILE:         ori? _str(ori.row[IDX.STOCK.FILE-1]||'') : '',
        CATEGORIA:    ori? _str(ori.row[IDX.STOCK.CATEGORIA-1]||categoria||'') : (categoria||'')
      };

      try{
        _ensureStockDelta(destino, producto, variante, +cant, true, metaPatch);
        _appendMov('Transferencia - Destino', destino, +cant, unitDest, moneda);
      }catch(e){ throw new Error('No se pudo actualizar destino: '+e.message); }

      try{
        const unitOri = unitAt(almacen);
        _ensureStockDelta(almacen, producto, variante, -cant, false, null);
        _appendMov('Transferencia - Origen', almacen, -cant, unitOri, moneda);
      }catch(e){
        try{
          _ensureStockDelta(destino, producto, variante, -cant, false, null);
          const unitRb = unitAt(destino);
          _appendMov('Rollback - Destino', destino, -cant, unitRb, moneda);
        }catch(_) {}
        throw new Error('No se pudo descontar en origen: '+e.message);
      }
      return { ok:true };
    }

    throw new Error('Tipo no soportado: '+rawTipo);
  }catch(err){
    return { ok:false, message: err.message||String(err) };
  }finally{
    try{ invalidateStockCache(); }catch(e){}
    try{ lock.releaseLock(); }catch(e){}
  }
}

/** ================== HISTORIAL / MÉTRICAS ================== **/
function getHistorialByDni(dni, limit) {
  try {
    const rows = _readRows(SHEPP.REGISTRO)
      .filter(r => _str(r[IDX.REG.DNI-1]) === _str(dni));

    const items = rows.map(r => {
      const producto = _str(r[IDX.REG.PRODUCTO-1]);
      const cargo = _str(r[IDX.REG.CARGO-1]);
      const operacion = _str(r[IDX.REG.OPERACION-1]);
      const fechaEntrega = r[IDX.REG.FECHA-1];
      
      // 🔧 Vida útil: primero desde REGISTRO, si no existe buscar en MATRIZ
      let vidaUtil = _num(r[IDX.REG.VIDA_UTIL_DIAS-1] || 0);
      
      if (vidaUtil === 0 && operacion === 'Entrega' && producto && cargo) {
        // Buscar en MATRIZ
        const reglas = getReglasCargo(producto, cargo);
        vidaUtil = _num(reglas.VIDA_UTIL_DIAS || 0);
      }

      // 🔧 Fecha de vencimiento: calcular si no existe
      let fechaVenc = r[IDX.REG.FECHA_VENCIMIENTO-1];
      
      // Si no hay fecha guardada Y es una entrega CON vida útil, calcular
      if (!fechaVenc && vidaUtil > 0 && operacion === 'Entrega' && fechaEntrega) {
        const d = new Date(fechaEntrega);
        d.setDate(d.getDate() + vidaUtil);
        fechaVenc = d; // guardar como objeto Date
      }

      return {
        FECHA:             _fmtDateOut(fechaEntrega),
        OPERACION:         operacion,
        ALMACEN:           _str(r[IDX.REG.ALMACEN-1]),
        PRODUCTO:          producto,
        VARIANTE:          _str(r[IDX.REG.VARIANTE-1] || ''),
        CANTIDAD:          _num(r[IDX.REG.CANTIDAD-1]),
        IMPORTE:           _num(r[IDX.REG.IMPORTE-1]),
        COSTO_UNITARIO:    _num(r[IDX.REG.COSTO_UNITARIO-1]),
        MONEDA:            _str(r[IDX.REG.MONEDA-1] || 'PEN'),
        DEVOLVIBLE:        _str(r[IDX.REG.DEVOLVIBLE-1] || ''),
        VIDA_UTIL_DIAS:    vidaUtil,                    // ✅ Enriquecido
        FECHA_VENCIMIENTO: _fmtDateOut(fechaVenc),      // ✅ Calculado si falta
        FIRMA_URL:         _str(r[IDX.REG.FIRMA_URL-1] || ''),
        OBS: _str(r[IDX.REG.OBS-1] || ''),
        ESTADO: _str(r[IDX.REG.ESTADO-1] || ''),
        FECHA_CONFIRMACION: _str(r[IDX.REG.FECHA_CONFIRMACION-1] || '')
      };
    })
    .sort((a, b) => new Date(b.FECHA) - new Date(a.FECHA));

    return (typeof limit === 'number' && limit > 0)
      ? items.slice(0, limit)
      : items;
      
  } catch (e) {
    console.error('Error en getHistorialByDni: ' + e.message);
    return [];
  }
}



function getCostoNetoByDni(dni){
  const reg=getHistorialByDni(dni);
  const tot={};
  reg.forEach(x=>{
    const k= x.PRODUCTO + '||' + (x.VARIANTE||'');
    if(!tot[k]) tot[k]={ MONTO:0, MONEDA:x.MONEDA||'PEN' };
    const sign=(x.OPERACION==='Devolución')?-1:+1;
    tot[k].MONTO += sign*(x.COSTO_UNITARIO*x.CANTIDAD);
  });
  return Object.keys(tot).map(k=>{
    const [p,v]=k.split('||');
    return { PRODUCTO:p, VARIANTE:v||'', MONTO_NETO: tot[k].MONTO, MONEDA: tot[k].MONEDA };
  });
}

/** ================== FIRMA ================== **/
function saveSignature(dataUrl) {
  if (!dataUrl) return { ok: false, message: 'Sin datos de firma' };

  // separar metadata y base64
  const parts = dataUrl.split(','),
        meta = parts[0],
        b64 = parts[1];

  const contentType = (meta.match(/data:(.*);base64/) || [])[1] || 'image/png';
  const bytes = Utilities.base64Decode(b64);
  const blob = Utilities.newBlob(bytes, contentType, 'firma_' + Date.now() + '.png');

  // 👉 Aquí pones el ID de tu carpeta de Drive
  const folder = DriveApp.getFolderById(FOLDER_IDEPP);

  // guardar archivo
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // URL directa (lh5.googleusercontent.com)
  const id = file.getId();
  const direct = 'https://lh5.googleusercontent.com/d/' + id;

  return { ok: true, url: direct, id };
}

/** ================== FLUJO TRABAJADOR: Pendientes / Confirmar / Rechazar ================== **/

/**
 * Obtiene todas las entregas pendientes de confirmación para un trabajador.
 * @param {string} dni - DNI del trabajador
 * @returns {Array} Lista de entregas pendientes
 */
function obtenerEntregasPendientes(dni){
  try{
    if(!dni) return [];
    const sh = _sh(SHEPP.REGISTRO);
    const data = sh.getDataRange().getValues();
    const pendientes = [];
    for(let i=1; i<data.length; i++){
      const r = data[i];
      if(_str(r[IDX.REG.DNI-1]) !== _str(dni)) continue;
      if(_str(r[IDX.REG.OPERACION-1]) !== 'Entrega') continue;
      if(_str(r[IDX.REG.ESTADO-1]) !== 'Pendiente') continue;
      pendientes.push({
        ID_REG: _str(r[IDX.REG.ID_REG-1]),
        FECHA: _fmtDateOut(r[IDX.REG.FECHA-1]),
        ALMACEN: _str(r[IDX.REG.ALMACEN-1]),
        PRODUCTO: _str(r[IDX.REG.PRODUCTO-1]),
        VARIANTE: _str(r[IDX.REG.VARIANTE-1]),
        CANTIDAD: _num(r[IDX.REG.CANTIDAD-1]),
        USUARIO: _str(r[IDX.REG.USUARIO-1]),
        OBS: _str(r[IDX.REG.OBS-1]),
        DEVOLVIBLE: _str(r[IDX.REG.DEVOLVIBLE-1]),
        ESTADO: 'Pendiente'
      });
    }
    return pendientes;
  }catch(e){
    console.error('Error en obtenerEntregasPendientes: ' + e.message);
    return [];
  }
}

/**
 * Cuenta entregas pendientes para un trabajador (para badge).
 * @param {string} dni - DNI del trabajador
 * @returns {number}
 */
function contarEntregasPendientes(dni){
  return obtenerEntregasPendientes(dni).length;
}

/**
 * Trabajador confirma la recepción de un EPP individual.
 * Cambia estado a "Confirmado", guarda firma, fecha y descuenta stock.
 * @param {string} idReg - ID del registro en REGISTRO
 * @param {string} firmaDataUrl - Base64 de la firma digital
 * @returns {Object} { ok, message }
 */
function confirmarEntregaEpp(idReg, firmaDataUrl){
  const lock = LockService.getScriptLock(); lock.tryLock(30000);
  try{
    if(!idReg) throw new Error('ID de registro requerido');
    if(!firmaDataUrl) throw new Error('Firma requerida para confirmar');

    // Subir firma a Drive
    const firmaRes = saveSignature(firmaDataUrl);
    if(!firmaRes || !firmaRes.ok) throw new Error('No se pudo guardar la firma');
    const firmaUrl = firmaRes.url;

    const ahora = new Date();
    const tz = Session.getScriptTimeZone();
    const fechaConf = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd HH:mm:ss');

    // Buscar y actualizar en REGISTRO
    const shReg = _sh(SHEPP.REGISTRO);
    const dataReg = shReg.getDataRange().getValues();
    let regRow = -1;
    let regData = null;
    for(let i=1; i<dataReg.length; i++){
      if(_str(dataReg[i][IDX.REG.ID_REG-1]) === _str(idReg)){
        regRow = i+1; // 1-based
        regData = dataReg[i];
        break;
      }
    }
    if(regRow < 0) throw new Error('Registro no encontrado: ' + idReg);
    if(_str(regData[IDX.REG.ESTADO-1]) !== 'Pendiente') throw new Error('Este EPP ya fue procesado');

    // Actualizar REGISTRO
    const totalColsReg = Math.max(shReg.getLastColumn(), IDX.REG.FECHA_CONFIRMACION);
    const regRange = shReg.getRange(regRow, 1, 1, totalColsReg);
    const regVals = regRange.getValues()[0];
    // Expandir si faltan columnas
    while(regVals.length < IDX.REG.FECHA_CONFIRMACION) regVals.push('');
    regVals[IDX.REG.FIRMA_URL-1] = firmaUrl;
    regVals[IDX.REG.ESTADO-1] = 'Confirmado';
    regVals[IDX.REG.FECHA_CONFIRMACION-1] = fechaConf;
    shReg.getRange(regRow, 1, 1, regVals.length).setValues([regVals]);

    // Buscar y actualizar en MOVIMIENTOS (mirror)
    const shMov = _sh(SHEPP.MOV);
    const dataMov = shMov.getDataRange().getValues();
    for(let i=1; i<dataMov.length; i++){
      const m = dataMov[i];
      if(_str(m[IDX.MOV.OPERACION-1]) === 'Entrega' &&
         _str(m[IDX.MOV.DNI-1]) === _str(regData[IDX.REG.DNI-1]) &&
         _str(m[IDX.MOV.PRODUCTO-1]) === _str(regData[IDX.REG.PRODUCTO-1]) &&
         _str(m[IDX.MOV.VARIANTE-1]||'') === _str(regData[IDX.REG.VARIANTE-1]||'') &&
         _str(m[IDX.MOV.FECHA-1]) === _str(regData[IDX.REG.FECHA-1]) &&
         _str(m[IDX.MOV.ESTADO-1]) === 'Pendiente'){
        const movRow = i+1;
        const totalColsMov = Math.max(shMov.getLastColumn(), IDX.MOV.FECHA_CONFIRMACION);
        const movRange = shMov.getRange(movRow, 1, 1, totalColsMov);
        const movVals = movRange.getValues()[0];
        while(movVals.length < IDX.MOV.FECHA_CONFIRMACION) movVals.push('');
        movVals[IDX.MOV.FIRMA_URL-1] = firmaUrl;
        movVals[IDX.MOV.ESTADO-1] = 'Confirmado';
        movVals[IDX.MOV.FECHA_CONFIRMACION-1] = fechaConf;
        shMov.getRange(movRow, 1, 1, movVals.length).setValues([movVals]);
        break;
      }
    }

    // Descontar stock AHORA
    const almacen = _str(regData[IDX.REG.ALMACEN-1]);
    const producto = _str(regData[IDX.REG.PRODUCTO-1]);
    const variante = _str(regData[IDX.REG.VARIANTE-1]);
    const cant = _num(regData[IDX.REG.CANTIDAD-1]);
    _ensureStockDelta(almacen, producto, variante, -cant, false, null);

    return { ok:true, message:'EPP confirmado correctamente' };
  }catch(err){
    return { ok:false, message: err.message||String(err) };
  }finally{
    try{ invalidateStockCache(); }catch(e){}
    try{ lock.releaseLock(); }catch(e){}
  }
}

/**
 * Trabajador rechaza un EPP.
 * Cambia estado a "Rechazado", guarda motivo, NO descuenta stock.
 * @param {string} idReg - ID del registro en REGISTRO
 * @param {string} motivo - Motivo del rechazo
 * @returns {Object} { ok, message }
 */
function rechazarEntregaEpp(idReg, motivo){
  const lock = LockService.getScriptLock(); lock.tryLock(30000);
  try{
    if(!idReg) throw new Error('ID de registro requerido');
    if(!motivo) throw new Error('Motivo de rechazo requerido');

    const ahora = new Date();
    const tz = Session.getScriptTimeZone();
    const fechaConf = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd HH:mm:ss');

    // Buscar y actualizar en REGISTRO
    const shReg = _sh(SHEPP.REGISTRO);
    const dataReg = shReg.getDataRange().getValues();
    let regRow = -1;
    let regData = null;
    for(let i=1; i<dataReg.length; i++){
      if(_str(dataReg[i][IDX.REG.ID_REG-1]) === _str(idReg)){
        regRow = i+1;
        regData = dataReg[i];
        break;
      }
    }
    if(regRow < 0) throw new Error('Registro no encontrado');
    if(_str(regData[IDX.REG.ESTADO-1]) !== 'Pendiente') throw new Error('Este EPP ya fue procesado');

    // Actualizar REGISTRO
    const totalColsReg = Math.max(shReg.getLastColumn(), IDX.REG.FECHA_CONFIRMACION);
    const regRange = shReg.getRange(regRow, 1, 1, totalColsReg);
    const regVals = regRange.getValues()[0];
    while(regVals.length < IDX.REG.FECHA_CONFIRMACION) regVals.push('');
    regVals[IDX.REG.ESTADO-1] = 'Rechazado';
    regVals[IDX.REG.FECHA_CONFIRMACION-1] = fechaConf;
    regVals[IDX.REG.OBS-1] = 'RECHAZADO: ' + _str(motivo) + (regVals[IDX.REG.OBS-1] ? ' | ' + regVals[IDX.REG.OBS-1] : '');
    shReg.getRange(regRow, 1, 1, regVals.length).setValues([regVals]);

    // Mirror en MOVIMIENTOS
    const shMov = _sh(SHEPP.MOV);
    const dataMov = shMov.getDataRange().getValues();
    for(let i=1; i<dataMov.length; i++){
      const m = dataMov[i];
      if(_str(m[IDX.MOV.OPERACION-1]) === 'Entrega' &&
         _str(m[IDX.MOV.DNI-1]) === _str(regData[IDX.REG.DNI-1]) &&
         _str(m[IDX.MOV.PRODUCTO-1]) === _str(regData[IDX.REG.PRODUCTO-1]) &&
         _str(m[IDX.MOV.VARIANTE-1]||'') === _str(regData[IDX.REG.VARIANTE-1]||'') &&
         _str(m[IDX.MOV.FECHA-1]) === _str(regData[IDX.REG.FECHA-1]) &&
         _str(m[IDX.MOV.ESTADO-1]) === 'Pendiente'){
        const movRow = i+1;
        const totalColsMov = Math.max(shMov.getLastColumn(), IDX.MOV.FECHA_CONFIRMACION);
        const movRange = shMov.getRange(movRow, 1, 1, totalColsMov);
        const movVals = movRange.getValues()[0];
        while(movVals.length < IDX.MOV.FECHA_CONFIRMACION) movVals.push('');
        movVals[IDX.MOV.ESTADO-1] = 'Rechazado';
        movVals[IDX.MOV.FECHA_CONFIRMACION-1] = fechaConf;
        movVals[IDX.MOV.OBS-1] = 'RECHAZADO: ' + _str(motivo) + (movVals[IDX.MOV.OBS-1] ? ' | ' + movVals[IDX.MOV.OBS-1] : '');
        shMov.getRange(movRow, 1, 1, movVals.length).setValues([movVals]);
        break;
      }
    }

    // NO se descuenta stock
    return { ok:true, message:'EPP rechazado. Motivo registrado.' };
  }catch(err){
    return { ok:false, message: err.message||String(err) };
  }finally{
    try{ invalidateStockCache(); }catch(e){}
    try{ lock.releaseLock(); }catch(e){}
  }
}

/**
 * Resumen de estados para el admin (conteo de Pendientes y Rechazados).
 * @param {string} almacen - Filtro opcional por almacén
 * @returns {Object} { pendientes, rechazados }
 */
function obtenerResumenEstados(almacen){
  try{
    const rows = _readRows(SHEPP.REGISTRO);
    let pendientes = 0, rechazados = 0;
    for(const r of rows){
      if(_str(r[IDX.REG.OPERACION-1]) !== 'Entrega') continue;
      if(almacen && _str(r[IDX.REG.ALMACEN-1]) !== _str(almacen)) continue;
      const estado = _str(r[IDX.REG.ESTADO-1]);
      if(estado === 'Pendiente') pendientes++;
      else if(estado === 'Rechazado') rechazados++;
    }
    return { pendientes, rechazados };
  }catch(e){
    return { pendientes:0, rechazados:0 };
  }
}

function crearProductoConOpcionalIngreso(p){
  const lock = LockService.getScriptLock(); lock.tryLock(30000);
  try{
    const producto = _str(p.producto);             if(!producto) throw new Error('Producto requerido');
    const almacen  = _str(p.almacen);              if(!almacen)  throw new Error('Almacén requerido');
    const variante = _str(p.variante||'');
    const crearIng = !!p.crearIngreso;
    const cant     = _num(p.cant||0);

    const patchMeta = {
      CATEGORIA:     _str(p.categoria||''),
      IMG:           _str(p.img||''),
      FILE:          _str(p.file||''),
      STOCK_MINIMO:  _num(p.minimo||0)
    };

    const costoUnitOpt = (p.costo===undefined || p.costo==='') ? null : _num(p.costo);
    const unit = _resolveUnitCost(almacen, producto, variante, costoUnitOpt);
    if (crearIng) patchMeta.PRECIO = unit;

    _ensureStockDelta(almacen, producto, variante, crearIng ? cant : 0, true, patchMeta);

    if (crearIng && cant>0){
      const hoy=_today();
      const m=_newRow(SHEPP.MOV);
      m[IDX.MOV.ID_MOV-1]          = _genId8();
      m[IDX.MOV.FECHA-1]           = hoy;
      m[IDX.MOV.OPERACION-1]       = 'Ingreso';
      m[IDX.MOV.ALMACEN-1]         = almacen;
      m[IDX.MOV.PRODUCTO-1]        = producto;
      m[IDX.MOV.VARIANTE-1]        = variante;
      m[IDX.MOV.CANTIDAD-1]        = cant;
      m[IDX.MOV.ID_PROV-1]         = _str(p.prov||'');
      m[IDX.MOV.MARCA-1]           = _str(p.marca||'');
      m[IDX.MOV.COSTO_UNITARIO-1]  = unit;
      m[IDX.MOV.MONEDA-1]          = _str(p.moneda||'PEN');
      m[IDX.MOV.IMPORTE-1]         = unit*cant;
      m[IDX.MOV.USUARIO-1]         = Session.getActiveUser().getEmail();
      m[IDX.MOV.OBS-1]             = _str(p.obs||'Ingreso inicial (nuevo producto)');
      _appendArray(SHEPP.MOV, m);
    }
    return { ok:true };
  }catch(err){
    return { ok:false, message: err.message||String(err) };
  }finally{
    try{ invalidateStockCache(); }catch(e){}
    try{ lock.releaseLock(); }catch(e){}
  }
}
function _resolveCategoria(almacen, producto, variante, categoriaOpt){
  const cand = _str(categoriaOpt||'');
  if (cand) return cand;

  const rows = _readRows(SHEPP.STOCK);

  // Exacta (almacén + producto + variante)
  let rec = rows.find(r =>
    _str(r[IDX.STOCK.ALMACEN-1])  === _str(almacen) &&
    _str(r[IDX.STOCK.PRODUCTO-1]) === _str(producto) &&
    _str(r[IDX.STOCK.VARIANTE-1]||'') === _str(variante||'')
  );
  if (rec && _str(rec[IDX.STOCK.CATEGORIA-1])) return _str(rec[IDX.STOCK.CATEGORIA-1]);

  // Por producto (cualquier variante del mismo producto en el almacén)
  rec = rows.find(r =>
    _str(r[IDX.STOCK.ALMACEN-1])  === _str(almacen) &&
    _str(r[IDX.STOCK.PRODUCTO-1]) === _str(producto) &&
    _str(r[IDX.STOCK.CATEGORIA-1])
  );
  if (rec) return _str(rec[IDX.STOCK.CATEGORIA-1]);

  // Fallback: MATRIZ GRID por producto base
  try{
    const grid = _readMatrizGrid_();
    const meta = grid && grid.byProduct ? grid.byProduct[_baseProducto(producto)] : null;
    if (meta && _str(meta.CATEGORIA||'')) return _str(meta.CATEGORIA);
  }catch(_) {}

  return '';
}
function getProductoMeta(almacenId, producto){
  const list = _readStockCache_();
  const rows = list.filter(x =>
    _str(x.ALMACEN)===_str(almacenId) &&
    _str(x.PRODUCTO)===_str(producto)
  );
  const firstWithFile = rows.find(r => _str(r.FILE));
  const firstWithImg  = rows.find(r => _str(r.IMG));
  const firstWithCat  = rows.find(r => _str(r.CATEGORIA));
  return {
    file: firstWithFile ? _str(firstWithFile.FILE) : '',
    img:  firstWithImg  ? _str(firstWithImg.IMG)   : '',
    categoria: firstWithCat ? _str(firstWithCat.CATEGORIA) : ''
  };
}

/** ================= INVALIDAR CACHÉS (stock + historial) ================= */
function invalidateStockCache(){
  try{
    const cache = CacheService.getDocumentCache();
    ['stock:all:v1','registro:all:v1','mov:all:v1'].forEach(k=>{
      try{ cache.remove(k); }catch(_){}
    });
  }catch(e){
    console.warn('No se pudo invalidar cachés:', e);
  }
}

/** (opcional) Alias interno para llamadas viejas desde el servidor */
function _invalidateStockCache_(){ return invalidateStockCache(); }



function crearProductoConIngresoOpcional(p){ return crearProductoConOpcionalIngreso(p); }


/****************************************************
 * CONFIG & UTILS
 ****************************************************/
const SHMATRIZ = {
  MATRIZ:     'MATRIZ',
  STOCK:      'STOCK',
  REGISTRO:   'REGISTRO',
  ALMACENES:  'ALMACENES'
};

// MATRIZ layout (basado en tu archivo):
// Fila 2: Req. capacitación
// Fila 3: Frec. inspección (días)
// Fila 4: Vida útil (días)
// Fila 5: Categoría
// Fila 6: Producto
// Fila 7+: Cargos en B, grilla de checks desde C
const MAT_FIRST_HEADER_ROW = 2; // fila 2
const MAT_HEADER_ROWS      = 5; // 5 filas (2..6)
const MAT_CARGOS_COL       = 2; // columna B
const MAT_FIRST_DATA_ROW   = 7; // fila 7
const MAT_FIRST_DATA_COL   = 3; // columna C

function transpose(m) {
  return m.length ? m[0].map((_, i) => m.map(r => r[i])) : [];
}

function _norm(v) {
  return String(v || '').split('|')[0].trim();
}

/****************************************************
 * MATRIZ – lectura/edición
 ****************************************************/

/**
 * Devuelve info del producto para un cargo específico:
 * - durabilidad (vida útil en días)
 * - previsto (✔︎ SI / ❌ NO)
 * - reqCapacitacion, frecInspeccion, categoria
 */
function getProductoInfo(nombreProducto, cargoEpp) {
  const sh = getSpreadsheetEPP().getSheetByName(SHMATRIZ.MATRIZ);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  const prodBuscado = _norm(nombreProducto);

  // 5 filas de encabezados por columna (C : last)
  const headerRange = sh.getRange(MAT_FIRST_HEADER_ROW, MAT_FIRST_DATA_COL, MAT_HEADER_ROWS, lastCol - (MAT_FIRST_DATA_COL - 1));
  const headers2to6 = headerRange.getValues(); // [ [cap...], [insp...], [vida...], [cat...], [prod...] ]
  const reqCapRow   = headers2to6[0];
  const frecInspRow = headers2to6[1];
  const vidaRow     = headers2to6[2];
  const catRow      = headers2to6[3];
  const prodRow     = headers2to6[4].map(_norm);

  const colIndex = prodRow.indexOf(prodBuscado);

  // Cargos en B7:B
  const cargosRng = sh.getRange(MAT_FIRST_DATA_ROW, MAT_CARGOS_COL, lastRow - (MAT_FIRST_DATA_ROW - 1), 1).getValues();
  const cargos = cargosRng.map(r => r[0]);
  const rowIndex = cargos.indexOf(cargoEpp);

  const durabilidad = (colIndex !== -1) ? (vidaRow[colIndex] || "-") : "-";
  const previsto = (rowIndex !== -1 && colIndex !== -1)
    ? (sh.getRange(MAT_FIRST_DATA_ROW + rowIndex, MAT_FIRST_DATA_COL + colIndex).getValue() === true ? "✔︎ SI" : "❌ NO")
    : "❌ NO";

  return {
    durabilidad,
    previsto,
    reqCapacitacion: (colIndex !== -1 ? (reqCapRow[colIndex]   || "") : ""),
    frecInspeccion:  (colIndex !== -1 ? (frecInspRow[colIndex] || "") : ""),
    categoria:       (colIndex !== -1 ? (catRow[colIndex]      || "") : "")
  };
}

/**
 * Construcción para matriz interactiva:
 * - encabezadosMatriz: por columna => [cap, insp, vida, cat, prod]
 * - checks: booleans de la grilla (C7..)
 * - nombresFila: cargos en B7:B
 * - opcionesProductos: productos únicos de STOCK!C
 */
function obtenerDatosMatrizInteractivaEpp() {
  const ss = getSpreadsheetEPP();
  const hojaMatriz = ss.getSheetByName(SHMATRIZ.MATRIZ);
  const hojaStock  = ss.getSheetByName(SHMATRIZ.STOCK);

  const lastRow = hojaMatriz.getLastRow();
  const lastCol = hojaMatriz.getLastColumn();

  // Encabezados (5 filas: 2..6) por columnas (C..last)
  const headerRows = hojaMatriz.getRange(MAT_FIRST_HEADER_ROW, MAT_FIRST_DATA_COL, MAT_HEADER_ROWS, lastCol - (MAT_FIRST_DATA_COL - 1)).getValues();
  const encabezadosMatriz = transpose(headerRows); // => [ [cap, insp, vida, cat, prod], ... ]

  // Nombres de fila (cargos) en B7:B
  const cargosVals = hojaMatriz.getRange(MAT_FIRST_DATA_ROW, MAT_CARGOS_COL, lastRow - (MAT_FIRST_DATA_ROW - 1), 1).getValues();
  const nombresFila = cargosVals.map(r => r[0]).filter(v => String(v || "").trim() !== "");

  // Checks desde C7, alineado a cantidad de cargos y columnas
  const totalCols = lastCol - (MAT_FIRST_DATA_COL - 1);
  const checks = hojaMatriz.getRange(MAT_FIRST_DATA_ROW, MAT_FIRST_DATA_COL, nombresFila.length, totalCols).getValues();

  // Opciones de PRODUCTO desde STOCK!C (únicas, limpias)
  const stockLastRow = hojaStock.getLastRow();
  const productos = (stockLastRow > 1)
    ? hojaStock.getRange(2, 3, stockLastRow - 1, 1).getValues().map(r => r[0])
    : [];
  const opcionesProductos = [...new Set(productos.filter(Boolean).map(_norm))];

  return { encabezadosMatriz, checks, nombresFila, opcionesProductos };
}

/**
 * Actualiza un checkbox de la grilla (fila/col son coordenadas absolutas de hoja)
 */
function actualizarCheckboxMatrizEpp(fila, col, val) {
  const hoja = getSpreadsheetEPP().getSheetByName(SHMATRIZ.MATRIZ);
  hoja.getRange(fila, col).setValue(val === true);
}

/**
 * Actualiza los encabezados de UNA columna (5 filas).
 * colInicio es índice absoluto de la columna en la hoja (empezando desde 1).
 * valores = [reqCap, frecInsp, vidaUtilDias, categoria, producto]
 */
function actualizarEncabezadoMatrizEpp(colInicio, valores) {
  const hoja = getSpreadsheetEPP().getSheetByName(SHMATRIZ.MATRIZ);

  // Normaliza valores (producto sin "|")
  const fixed = [
    String(valores[0] || '').trim(),           // Req. capacitación
    String(valores[1] || '').trim(),           // Frec. inspección (días)
    String(valores[2] || '').trim(),           // Vida útil (días)
    String(valores[3] || '').trim(),           // Categoría
    _norm(valores[4])                          // Producto
  ];

  hoja.getRange(MAT_FIRST_HEADER_ROW, colInicio, MAT_HEADER_ROWS, 1)
      .setValues(fixed.map(v => [v]));
}

/**
 * Inserta una nueva columna al final (después de la última usada)
 */
function agregarNuevaColumnaMatrizEpp() {
  const hoja = getSpreadsheetEPP().getSheetByName(SHMATRIZ.MATRIZ);
  const lastCol = hoja.getLastColumn();
  hoja.insertColumnAfter(lastCol);
  // (Opcional) podrías inicializar los 5 encabezados vacíos aquí si lo deseas.
}

/**
 * Elimina una columna de la matriz dado su índice relativo (0-based) desde C.
 * Ej: indiceCol=0 => borra la columna C; 1 => D; etc.
 */
function eliminarColumnaMatrizEpp(indiceCol) {
  const hoja = getSpreadsheetEPP().getSheetByName(SHMATRIZ.MATRIZ);
  const col = MAT_FIRST_DATA_COL + indiceCol; // desde C
  hoja.deleteColumn(col);
}

/****************************************************
 * OTRAS FUNCIONES EXISTENTES (se mantienen)
 ****************************************************/

/**
 * Últimas 2 entregas por ID y Producto (REGISTRO)
 * Usa B: fecha, C: ID, J: operación, K: producto
 */
function getUltimasEntregas(id, producto) {
  const hoja = getSpreadsheetEPP().getSheetByName(SHMATRIZ.REGISTRO);
  const lastRow = hoja.getLastRow();
  if (lastRow < 2) return [];

  // Leer B2:K (10 columnas)
  const datos = hoja.getRange(2, 2, lastRow - 1, 10).getValues(); // B..K

  const resultados = datos
    .filter(row => row[1] == id && row[9] == producto) // C=ID (idx1), K=Producto (idx9)
    .sort((a, b) => new Date(b[0]) - new Date(a[0]))   // B=Fecha (idx0)
    .slice(0, 2)
    .map(row => ({
      fecha: Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
      operacion: row[8] // J = Operación (idx8)
    }));

  return resultados;
}

/**
 * Lista de almacenes (columna B de ALMACENES)
 */
function getAlmacenes() {
  const sheet = getSpreadsheetEPP().getSheetByName(SHMATRIZ.ALMACENES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B2:B
  return data.flat().filter(v => String(v || '').trim() !== '');
}

/** ================== LISTAS PARA MODAL MATRIZ ================== **/
function getListasMatrizEpp() {
  const shLista = _sh('LISTA');
  const shStock = _sh(SHEPP.STOCK);

  const reqCap = shLista.getRange('B2:B').getValues()
    .map(r => _str(r[0]))
    .filter(Boolean)
    .filter((v, i, a) => a.indexOf(v) === i)
    .sort();

  const categorias = shStock.getRange('E2:E').getValues()
    .map(r => _str(r[0]))
    .filter(Boolean)
    .filter((v, i, a) => a.indexOf(v) === i)
    .sort();

  return { reqCap, categorias };
}




function obtenerMaestroCompleto() {
  try {
    const ss = getSpreadsheetEPP();
    const shReg = ss.getSheetByName(SHEPP.REGISTRO);
    
    if (!shReg) {
      throw new Error('No se encontró la hoja REGISTRO');
    }
    
    const data = shReg.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // 🔍 Índices de columnas
    const idx = {
      FECHA: headers.indexOf('FECHA'),
      OPERACION: headers.indexOf('OPERACION'),
      DNI: headers.indexOf('DNI'),
      NOMBRES: headers.indexOf('NOMBRES'),
      CARGO: headers.indexOf('CARGO'),
      PRODUCTO: headers.indexOf('PRODUCTO'),
      VARIANTE: headers.indexOf('VARIANTE'),
      FECHA_VENCIMIENTO: headers.indexOf('FECHA_VENCIMIENTO'),
      VIDA_UTIL_DIAS: headers.indexOf('VIDA_UTIL_DIAS')
    };
    
    // 📊 Procesar solo entregas
    const entregas = rows
      .filter(r => _str(r[idx.OPERACION]) === 'Entrega')
      .map(r => {
        const fechaVenc = r[idx.FECHA_VENCIMIENTO];
        const diasVida = Number(r[idx.VIDA_UTIL_DIAS] || 0);
        
        // 🔧 Calcular estado
        let estado = 'SIN ASIGNAR';
        let clase = 'sin-asignar';
        let diasRestantes = 0;
        
        if (fechaVenc && diasVida > 0) {
          const hoy = new Date();
          hoy.setHours(0, 0, 0, 0);
          
          const venc = new Date(fechaVenc);
          venc.setHours(0, 0, 0, 0);
          
          const diffMs = venc - hoy;
          diasRestantes = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
          
          if (diasRestantes < 0) {
            estado = 'VENCIDO';
            clase = 'st-ENTR';
          } else if (diasRestantes <= 15) {
            estado = 'POR VENCER';
            clase = 'st-NOTIF';
          } else {
            estado = 'VIGENTE';
            clase = 'st-OK';
          }
        }
        
        return {
          fecha: _fmtDateOut(r[idx.FECHA]),
          trabajador: _str(r[idx.NOMBRES]),
          cargo: _str(r[idx.CARGO]),
          producto: _str(r[idx.PRODUCTO]),
          variante: _str(r[idx.VARIANTE]),
          vencimiento: _fmtDateOut(fechaVenc),
          estado: estado,
          clase: clase,
          diasRestantes: diasRestantes
        };
      });
    
    Logger.log(`✅ Maestro completo: ${entregas.length} registros`);
    return entregas;
    
  } catch(e) {
    Logger.log('❌ Error en obtenerMaestroCompleto: ' + e.message);
    throw new Error('Error al obtener maestro: ' + e.message);
  }
}

/****************************************************
 * EXPORTAR PDF DE PENDIENTES
 ****************************************************/
function generarPDFPendientes(datos) {
  try {
    if (!Array.isArray(datos) || datos.length === 0) {
      throw new Error('No hay datos para exportar');
    }

    // Crear nuevo documento PDF en Drive
    const folder = DriveApp.getFolderById(FOLDER_IDEPP);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const fileName = `EPP_Pendientes_${timestamp}`;
    
    // Crear hoja temporal para exportar
    const ss = SpreadsheetApp.create(fileName);
    const sheet = ss.getActiveSheet();
    sheet.setName('Pendientes');
    
    // Encabezados con formato
    const headers = [
      'Trabajador', 'Cargo', 'Producto', 'Variante', 
      'Fecha Vencimiento', 'Estado', 'Días Restantes'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold')
      .setBackground('#28a745').setFontColor('#ffffff');
    
    // Datos
    const rows = datos.map(d => [
      d.trabajador || '',
      d.cargo || '',
      d.producto || '',
      d.variante || '',
      d.vencimiento || '',
      d.estado || '',
      d.diasRestantes || 0
    ]);
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      
      // Aplicar colores según estado
      for (let i = 0; i < rows.length; i++) {
        const estado = rows[i][5]; // columna Estado
        const rowNum = i + 2;
        let color = '#ffffff';
        
        if (estado === 'VENCIDO') color = '#f8d7da';
        else if (estado === 'POR VENCER') color = '#fff3cd';
        
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground(color);
      }
    }
    
    // Ajustar columnas
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    // Convertir a PDF
    const file = DriveApp.getFileById(ss.getId());
    const blob = file.getAs('application/pdf');
    const pdfFile = folder.createFile(blob.setName(fileName + '.pdf'));
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Eliminar hoja temporal
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    
    return pdfFile.getUrl();
    
  } catch(e) {
    console.error('Error en generarPDFPendientes:', e);
    throw new Error('Error al generar PDF: ' + e.message);
  }
}

/****************************************************
 * GENERAR LISTA DE COMPRA
 ****************************************************/
function generarListaCompra(agrupado) {
  try {
    if (!agrupado || typeof agrupado !== 'object') {
      throw new Error('Datos inválidos para lista de compra');
    }

    const folder = DriveApp.getFolderById(FOLDER_IDEPP);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const fileName = `EPP_ListaCompra_${timestamp}`;
    
    // Crear hoja temporal
    const ss = SpreadsheetApp.create(fileName);
    const sheet = ss.getActiveSheet();
    sheet.setName('Lista de Compra');
    
    // Título
    sheet.getRange(1, 1, 1, 4).merge().setValue('LISTA DE COMPRA - EPP VENCIDOS')
      .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground('#dc3545').setFontColor('#ffffff');
    
    // Fecha
    sheet.getRange(2, 1, 1, 4).merge()
      .setValue('Generado: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))
      .setFontStyle('italic').setHorizontalAlignment('center');
    
    // Encabezados
    const headers = ['Producto', 'Cantidad Requerida', 'Trabajadores Afectados', 'Detalles'];
    sheet.getRange(4, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#ffc107').setFontColor('#212529');
    
    // Datos
    let rowNum = 5;
    for (const key in agrupado) {
      const item = agrupado[key];
      
      // Detalles de trabajadores
      const detalles = item.trabajadores.map(t => 
        `${t.nombre} (${t.variante || 'N/A'}) - Vence: ${t.vencimiento}`
      ).join('\n');
      
      sheet.getRange(rowNum, 1).setValue(item.producto);
      sheet.getRange(rowNum, 2).setValue(item.cantidad).setHorizontalAlignment('center');
      sheet.getRange(rowNum, 3).setValue(item.trabajadores.length).setHorizontalAlignment('center');
      sheet.getRange(rowNum, 4).setValue(detalles).setWrap(true);
      
      rowNum++;
    }
    
    // Ajustar columnas
    sheet.setColumnWidth(1, 200); // Producto
    sheet.setColumnWidth(2, 150); // Cantidad
    sheet.setColumnWidth(3, 180); // Trabajadores
    sheet.setColumnWidth(4, 400); // Detalles
    
    // Bordes
    const lastRow = rowNum - 1;
    sheet.getRange(4, 1, lastRow - 3, 4).setBorder(
      true, true, true, true, true, true,
      '#000000', SpreadsheetApp.BorderStyle.SOLID
    );
    
    // Resumen
    const totalProductos = Object.keys(agrupado).length;
    const totalUnidades = Object.values(agrupado).reduce((sum, item) => sum + item.cantidad, 0);
    
    sheet.getRange(rowNum + 1, 1).setValue('RESUMEN:').setFontWeight('bold');
    sheet.getRange(rowNum + 2, 1).setValue(`Total de productos diferentes: ${totalProductos}`);
    sheet.getRange(rowNum + 3, 1).setValue(`Total de unidades a comprar: ${totalUnidades}`);
    
    // Convertir a PDF
    const file = DriveApp.getFileById(ss.getId());
    const blob = file.getAs('application/pdf');
    const pdfFile = folder.createFile(blob.setName(fileName + '.pdf'));
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Eliminar hoja temporal
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    
    return pdfFile.getUrl();
    
  } catch(e) {
    console.error('Error en generarListaCompra:', e);
    throw new Error('Error al generar lista: ' + e.message);
  }
}

/****************************************************
 * OBTENER IMÁGENES DE EPP (Optimizado)
 ****************************************************/
function getImagenesEPP(productos) {
  try {
    if (!Array.isArray(productos)) return {};
    
    const cache = CacheService.getDocumentCache();
    const resultado = {};
    const PLACEHOLDER = 'https://lh5.googleusercontent.com/d/1n38-t1dDdj56TqOjv8c4SAAcUQcA0U72';
    
    productos.forEach(prod => {
      const cacheKey = 'img:' + prod;
      const cached = cache.get(cacheKey);
      
      if (cached) {
        resultado[prod] = cached;
      } else {
        // Buscar en STOCK
        const stockRows = _readStockCache_();
        const found = stockRows.find(r => _str(r.PRODUCTO) === _str(prod) && _str(r.IMG));
        
        const img = found ? _str(found.IMG) : PLACEHOLDER;
        resultado[prod] = img;
        
        // Cachear por 10 minutos
        try { cache.put(cacheKey, img, 600); } catch(e) {}
      }
    });
    
    return resultado;
    
  } catch(e) {
    console.error('Error en getImagenesEPP:', e);
    return {};
  }
}
/**
 * 🔧 SCRIPT DE CORRECCIÓN MANUAL
 * Recalcula todas las fechas de vencimiento incorrectas
 * Ejecutar UNA SOLA VEZ después de corregir registrarEntrega()
 */
function recalcularFechasVencimiento() {
  try {
    const ss = getSpreadsheetEPP();
    const sh = ss.getSheetByName(SHEPP.REGISTRO);
    const data = sh.getDataRange().getValues();
    
    let corregidos = 0;
    
    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      
      const operacion = _str(fila[IDX.REG.OPERACION - 1]);
      if (operacion !== 'Entrega') continue;
      
      const fechaEntrega = fila[IDX.REG.FECHA - 1];
      const vidaUtil = _num(fila[IDX.REG.VIDA_UTIL_DIAS - 1] || 0);
      
      if (!fechaEntrega || vidaUtil <= 0) continue;
      
      // Calcular correctamente
      const d = new Date(fechaEntrega);
      d.setDate(d.getDate() + vidaUtil);
      const nuevaFecha = Utilities.formatDate(d, 'GMT-5', 'yyyy-MM-dd');
      
      const fechaActual = _str(fila[IDX.REG.FECHA_VENCIMIENTO - 1]);
      
      if (fechaActual !== nuevaFecha) {
        sh.getRange(i + 1, IDX.REG.FECHA_VENCIMIENTO).setValue(nuevaFecha);
        
        const producto = _str(fila[IDX.REG.PRODUCTO - 1]);
        const dni = _str(fila[IDX.REG.DNI - 1]);
        Logger.log(`✅ Fila ${i+1}: ${producto} (DNI: ${dni}) | ${fechaActual} → ${nuevaFecha}`);
        corregidos++;
      }
    }
    
    Logger.log(`\n🎉 Proceso completado: ${corregidos} registros corregidos de ${data.length - 1} entregas`);
    invalidateStockCache();
    
    return { ok: true, corregidos, total: data.length - 1 };
    
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    return { ok: false, message: e.message };
  }
}
/**
 * Busca la última entrega del mismo producto/variante para un DNI
 * @param {string} dni - DNI del trabajador
 * @param {string} producto - Producto base
 * @param {string} variante - Variante (opcional)
 * @returns {Object|null} { FECHA, FECHA_VENCIMIENTO } o null
 */
function buscarUltimaEntrega(dni, producto, variante){
  try {
    const rows = _readRows(SHEPP.REGISTRO);
    const entregas = rows
      .filter(r => 
        _str(r[IDX.REG.DNI-1]) === _str(dni) &&
        _str(r[IDX.REG.PRODUCTO-1]) === _str(producto) &&
        _str(r[IDX.REG.VARIANTE-1]||'') === _str(variante||'') &&
        _str(r[IDX.REG.OPERACION-1]) === 'Entrega'
      )
      .sort((a,b) => {
        const dateA = new Date(a[IDX.REG.FECHA-1]);
        const dateB = new Date(b[IDX.REG.FECHA-1]);
        return dateB - dateA; // Más reciente primero
      });
    
    if (!entregas.length) return null;
    
    const r = entregas[0];
    return {
      FECHA: r[IDX.REG.FECHA-1],
      FECHA_VENCIMIENTO: r[IDX.REG.FECHA_VENCIMIENTO-1]
    };
  } catch(e) {
    Logger.log('Error en buscarUltimaEntrega: ' + e);
    return null;
  }
}

/**
 * 🆕 AGREGA OBSERVACIONES DE RETRASO/A TIEMPO (VERSIÓN MEJORADA)
 * Actualiza columna OBS en REGISTRO y MOVIMIENTOS
 * NO modifica fechas de vencimiento
 */
function agregarObservacionesRetraso() {
  try {
    const ss = getSpreadsheetEPP();
    const shReg = ss.getSheetByName(SHEPP.REGISTRO);
    const shMov = ss.getSheetByName(SHEPP.MOV);
    
    if (!shReg || !shMov) {
      throw new Error('No se encontraron las hojas REGISTRO o MOVIMIENTOS');
    }
    
    const dataReg = shReg.getDataRange().getValues();
    const dataMov = shMov.getDataRange().getValues();
    
    let actualizadosReg = 0;
    let actualizadosMov = 0;
    let primerasEntregas = 0;
    
    // ========== PROCESAR REGISTRO ==========
    Logger.log('🔄 Procesando hoja REGISTRO...');
    Logger.log(`Total de filas: ${dataReg.length - 1}`);
    
    for (let i = 1; i < dataReg.length; i++) {
      const fila = dataReg[i];
      
      // Validar que la fila tenga datos
      if (!fila || fila.length < IDX.REG.OBS) {
        Logger.log(`⚠️ Fila ${i+1}: Sin datos suficientes`);
        continue;
      }
      
      const operacion = _str(fila[IDX.REG.OPERACION - 1]);
      if (operacion !== 'Entrega') continue;
      
      const dni = _str(fila[IDX.REG.DNI - 1]);
      const producto = _str(fila[IDX.REG.PRODUCTO - 1]);
      const variante = _str(fila[IDX.REG.VARIANTE - 1] || '');
      
      // Validar fecha
      const fechaEntregaRaw = fila[IDX.REG.FECHA - 1];
      if (!fechaEntregaRaw) {
        Logger.log(`⚠️ Fila ${i+1}: Sin fecha de entrega`);
        continue;
      }
      
      const fechaEntrega = new Date(fechaEntregaRaw);
      if (isNaN(fechaEntrega.getTime())) {
        Logger.log(`⚠️ Fila ${i+1}: Fecha inválida`);
        continue;
      }
      
      const obsActual = _str(fila[IDX.REG.OBS - 1]);
      
      // 🔍 Buscar entrega anterior del mismo producto
      const ultimaEntrega = buscarUltimaEntregaAnterior(dni, producto, variante, fechaEntrega, dataReg);
      
      let nuevaObs = '';
      
      if (!ultimaEntrega || !ultimaEntrega.FECHA_VENCIMIENTO) {
        // 🆕 ES LA PRIMERA ENTREGA
        const primeraMsg = '🆕 Primera entrega registrada.';
        
        if (!obsActual.includes('Primera entrega') && !obsActual.includes('🆕')) {
          nuevaObs = obsActual ? `${primeraMsg} ${obsActual}` : primeraMsg;
          primerasEntregas++;
        }
      } else {
        // 🔄 HAY ENTREGA ANTERIOR - COMPARAR FECHAS
        const vencPrev = new Date(ultimaEntrega.FECHA_VENCIMIENTO);
        
        if (isNaN(vencPrev.getTime())) {
          Logger.log(`⚠️ Fila ${i+1}: Fecha de vencimiento anterior inválida`);
          continue;
        }
        
        if (fechaEntrega > vencPrev) {
          // 🚨 RETRASO
          const diffMs = fechaEntrega - vencPrev;
          const diasRetraso = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
          const retrasoMsg = `⚠️ Reemplazo con retraso de ${diasRetraso} día${diasRetraso===1?'':'s'}.`;
          
          if (!obsActual.includes('Reemplazo con retraso') && !obsActual.includes('⚠️')) {
            nuevaObs = obsActual ? `${retrasoMsg} ${obsActual}` : retrasoMsg;
          }
        } else {
          // ✅ A TIEMPO
          const aTiempoMsg = '✅ Reemplazo a tiempo.';
          
          if (!obsActual.includes('Reemplazo a tiempo') && !obsActual.includes('✅')) {
            nuevaObs = obsActual ? `${aTiempoMsg} ${obsActual}` : aTiempoMsg;
          }
        }
      }
      
      // Actualizar solo si hay cambio
      if (nuevaObs) {
        shReg.getRange(i + 1, IDX.REG.OBS).setValue(nuevaObs);
        Logger.log(`✅ REG Fila ${i+1}: ${producto} (DNI: ${dni}) | ${nuevaObs.substring(0, 60)}...`);
        actualizadosReg++;
      }
    }
    
    // ========== PROCESAR MOVIMIENTOS ==========
    Logger.log('\n🔄 Procesando hoja MOVIMIENTOS...');
    Logger.log(`Total de filas: ${dataMov.length - 1}`);
    
    for (let i = 1; i < dataMov.length; i++) {
      const fila = dataMov[i];
      
      // Validar que la fila tenga datos
      if (!fila || fila.length < IDX.MOV.OBS) continue;
      
      const operacion = _str(fila[IDX.MOV.OPERACION - 1]);
      if (operacion !== 'Entrega') continue;
      
      const dni = _str(fila[IDX.MOV.DNI - 1]);
      const producto = _str(fila[IDX.MOV.PRODUCTO - 1]);
      const variante = _str(fila[IDX.MOV.VARIANTE - 1] || '');
      
      // Validar fecha
      const fechaEntregaRaw = fila[IDX.MOV.FECHA - 1];
      if (!fechaEntregaRaw) continue;
      
      const fechaEntrega = new Date(fechaEntregaRaw);
      if (isNaN(fechaEntrega.getTime())) continue;
      
      const obsActual = _str(fila[IDX.MOV.OBS - 1]);
      
      // 🔍 Buscar entrega anterior (en REGISTRO, no en MOV)
      const ultimaEntrega = buscarUltimaEntregaAnterior(dni, producto, variante, fechaEntrega, dataReg);
      
      let nuevaObs = '';
      
      if (!ultimaEntrega || !ultimaEntrega.FECHA_VENCIMIENTO) {
        // 🆕 ES LA PRIMERA ENTREGA
        const primeraMsg = '🆕 Primera entrega registrada.';
        
        if (!obsActual.includes('Primera entrega') && !obsActual.includes('🆕')) {
          nuevaObs = obsActual ? `${primeraMsg} ${obsActual}` : primeraMsg;
        }
      } else {
        // 🔄 HAY ENTREGA ANTERIOR
        const vencPrev = new Date(ultimaEntrega.FECHA_VENCIMIENTO);
        if (isNaN(vencPrev.getTime())) continue;
        
        if (fechaEntrega > vencPrev) {
          // 🚨 RETRASO
          const diffMs = fechaEntrega - vencPrev;
          const diasRetraso = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
          const retrasoMsg = `⚠️ Reemplazo con retraso de ${diasRetraso} día${diasRetraso===1?'':'s'}.`;
          
          if (!obsActual.includes('Reemplazo con retraso') && !obsActual.includes('⚠️')) {
            nuevaObs = obsActual ? `${retrasoMsg} ${obsActual}` : retrasoMsg;
          }
        } else {
          // ✅ A TIEMPO
          const aTiempoMsg = '✅ Reemplazo a tiempo.';
          
          if (!obsActual.includes('Reemplazo a tiempo') && !obsActual.includes('✅')) {
            nuevaObs = obsActual ? `${aTiempoMsg} ${obsActual}` : aTiempoMsg;
          }
        }
      }
      
      // Actualizar solo si hay cambio
      if (nuevaObs) {
        shMov.getRange(i + 1, IDX.MOV.OBS).setValue(nuevaObs);
        Logger.log(`✅ MOV Fila ${i+1}: ${producto} (DNI: ${dni}) | ${nuevaObs.substring(0, 60)}...`);
        actualizadosMov++;
      }
    }
    
    Logger.log(`\n🎉 Proceso completado:`);
    Logger.log(`   REGISTRO: ${actualizadosReg} observaciones agregadas (${primerasEntregas} primeras entregas)`);
    Logger.log(`   MOVIMIENTOS: ${actualizadosMov} observaciones agregadas`);
    Logger.log(`   Total entregas en REGISTRO: ${dataReg.length - 1}`);
    Logger.log(`   Total entregas en MOVIMIENTOS: ${dataMov.length - 1}`);
    
    invalidateStockCache();
    
    return { 
      ok: true, 
      registro: actualizadosReg, 
      movimientos: actualizadosMov,
      primerasEntregas: primerasEntregas,
      totalReg: dataReg.length - 1,
      totalMov: dataMov.length - 1
    };
    
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    return { ok: false, message: e.message };
  }
}

/**
 * 🔍 BUSCA LA ENTREGA ANTERIOR AL MISMO PRODUCTO/VARIANTE
 * (Excluye la fila actual usando la fecha)
 */
function buscarUltimaEntregaAnterior(dni, producto, variante, fechaActual, dataReg) {
  try {
    // Validar parámetros
    if (!dataReg || !Array.isArray(dataReg) || dataReg.length < 2) {
      return null;
    }
    
    if (!dni || !producto || !fechaActual) {
      return null;
    }
    
    const entregas = [];
    
    for (let i = 1; i < dataReg.length; i++) {
      const r = dataReg[i];
      
      // Validar que la fila tenga suficientes columnas
      if (!r || r.length < IDX.REG.FECHA_VENCIMIENTO) continue;
      
      // Validar DNI
      if (_str(r[IDX.REG.DNI-1]) !== _str(dni)) continue;
      
      // Validar Producto
      if (_str(r[IDX.REG.PRODUCTO-1]) !== _str(producto)) continue;
      
      // Validar Variante
      if (_str(r[IDX.REG.VARIANTE-1]||'') !== _str(variante||'')) continue;
      
      // Validar Operación
      if (_str(r[IDX.REG.OPERACION-1]) !== 'Entrega') continue;
      
      // Validar Fecha
      const fechaRegRaw = r[IDX.REG.FECHA-1];
      if (!fechaRegRaw) continue;
      
      const fechaReg = new Date(fechaRegRaw);
      if (isNaN(fechaReg.getTime())) continue;
      
      // 🚫 Excluir la fila actual (misma fecha Y mismo producto)
      if (fechaReg.getTime() === fechaActual.getTime()) continue;
      
      // ✅ Solo entregas ANTERIORES
      if (fechaReg < fechaActual) {
        const fechaVenc = r[IDX.REG.FECHA_VENCIMIENTO-1];
        if (fechaVenc) {
          entregas.push({
            FECHA: fechaReg,
            FECHA_VENCIMIENTO: fechaVenc
          });
        }
      }
    }
    
    // Ordenar por fecha DESC (más reciente primero)
    entregas.sort((a,b) => b.FECHA - a.FECHA);
    
    return entregas[0] || null;
    
  } catch (e) {
    Logger.log('❌ Error en buscarUltimaEntregaAnterior: ' + e.message);
    return null;
  }
}