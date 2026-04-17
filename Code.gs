const SPREADSHEET_ID = "1sYoKHijD6Pfyd-k8SolYaVUSKIr8BiJqZ3a367lTtQ8";
const HOJA_PRODUCTOS = "Productos";
const HOJA_MOVIMIENTOS = "Movimientos";
const HOJA_UNIDADES = "Unidades";
const HOJA_GRUPOS = "Grupos";
const HOJA_CLIENTES = "CLIENTES";
const HOJA_CAJA = "CAJA_REGISTRADORA";
const HOJA_KITS = "KITS";
const HOJA_KIT_COMPONENTES = "KIT_COMPONENTES";
const HOJA_PARAMETROS = "PARAMETROS";

const TIPOS_MOVIMIENTO = {
  INGRESO: "INGRESO",
  SALIDA: "SALIDA",
  AJUSTE_POSITIVO: "AJUSTE_POSITIVO",
  AJUSTE_NEGATIVO: "AJUSTE_NEGATIVO",
  AJUSTE: "AJUSTE",
  ENTRADA_MANTTO: "ENTRADA_X_MANTENIMIENTO",
  SALIDA_MANTTO: "SALIDA_X_MANTENIMIENTO",
  SALIDA_SCRAP: "SALIDA_X_SCRAP",
  ENTRADA_COMPRA_PROV: "ENTRADA_X_COMPRA_PROVEEDOR",
  DEV_COMPRA_PROV: "DEVOLUCION_X_COMPRA_PROVEEDOR",
  SALIDA_VENTA: "SALIDA_X_VENTA",
  DEV_VENTA: "DEVOLUCION_X_VENTA",
  ENTRADA_MANTTO_CLI: "ENTRADA_X_MANTENIMIENTO_CLIENTE",
  SALIDA_MANTTO_CLI: "SALIDA_X_MANTENIMIENTO_CLIENTE",
  SALIDA_RENTA: "SALIDA_X_RENTA",
  ENTRADA_RENTA: "ENTRADA_X_RENTA",
  DEVOLUCION_VENTA: "ENTRADA_X_DEVOLUCION"
};

const TIPOS_ENTRADA = new Set([
  TIPOS_MOVIMIENTO.INGRESO,
  TIPOS_MOVIMIENTO.AJUSTE_POSITIVO,
  TIPOS_MOVIMIENTO.AJUSTE,
  TIPOS_MOVIMIENTO.ENTRADA_MANTTO,
  TIPOS_MOVIMIENTO.ENTRADA_COMPRA_PROV,
  TIPOS_MOVIMIENTO.DEV_VENTA,
  TIPOS_MOVIMIENTO.ENTRADA_MANTTO_CLI,
  TIPOS_MOVIMIENTO.ENTRADA_RENTA,
  TIPOS_MOVIMIENTO.DEVOLUCION_VENTA
]);

const TIPOS_SALIDA = new Set([
  TIPOS_MOVIMIENTO.SALIDA,
  TIPOS_MOVIMIENTO.AJUSTE_NEGATIVO,
  TIPOS_MOVIMIENTO.SALIDA_MANTTO,
  TIPOS_MOVIMIENTO.SALIDA_SCRAP,
  TIPOS_MOVIMIENTO.DEV_COMPRA_PROV,
  TIPOS_MOVIMIENTO.SALIDA_VENTA,
  TIPOS_MOVIMIENTO.SALIDA_MANTTO_CLI,
  TIPOS_MOVIMIENTO.SALIDA_RENTA
]);

const TIPOS_CORRECCION_MAP = {
  INGRESOXRENTA: TIPOS_MOVIMIENTO.ENTRADA_RENTA
};

function normalizarTipoMovimiento(tipo) {
  if (!tipo) return '';
  return String(tipo).trim().toUpperCase().replace(/\s+/g, '_');
}
function esTipoRentaFlexible(tipo) {
  const t = normalizarTipoMovimiento(tipo);
  return t === TIPOS_MOVIMIENTO.SALIDA_RENTA || t.indexOf('RENTA') !== -1;
}
function corregirTipoMovimientoValor(valor) {
  const norm = normalizarTipoMovimiento(valor);
  return TIPOS_CORRECCION_MAP[norm] || norm;
}
function esEntrada(tipo) { return TIPOS_ENTRADA.has(tipo); }
function esSalida(tipo) { return TIPOS_SALIDA.has(tipo); }

/** WEB APP */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Sistema de Control de Inventario");
}

/* ------------------- CLIENTES ------------------- */
function generarClaveCliente() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HOJA_CLIENTES);
  if (!sheet) sheet = ss.insertSheet(HOJA_CLIENTES);
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const val = data[i][0];
    if (val && typeof val === 'string' && val.toUpperCase().startsWith('CLI')) {
      const num = parseInt(val.replace(/[^0-9]/g, ''), 10);
      if (!isNaN(num)) maxNum = Math.max(maxNum, num);
    }
  }
  return 'CLI' + String(maxNum + 1).padStart(3, '0');
}
function obtenerClaveCliente() { return generarClaveCliente(); }

function registrarCliente(cliente) {
  try {
    if (!cliente || !cliente.nombre) throw new Error('Nombre es obligatorio');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(HOJA_CLIENTES);
    if (!sheet) sheet = ss.insertSheet(HOJA_CLIENTES);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 7).setValues([["Clave","Nombre","Teléfono","Email","Dirección","Estado","Fecha Creación"]]);
    }
    const data = sheet.getDataRange().getValues();
    const clave = (cliente.clave || '').toString().trim().toUpperCase() || generarClaveCliente();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toUpperCase() === clave) return 'Ya existe un cliente con esa clave';
    }
    sheet.appendRow([clave, cliente.nombre.trim(), cliente.telefono || '', cliente.email || '', cliente.direccion || '', 'ACTIVO', new Date()]);
    return 'Cliente registrado correctamente';
  } catch (err) { return 'Error: ' + err.message; }
}

function actualizarCliente(cliente) {
  try {
    if (!cliente || !cliente.clave) throw new Error('Clave requerida');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_CLIENTES);
    const data = sheet.getDataRange().getValues();
    const clave = String(cliente.clave).toUpperCase();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).toUpperCase() === clave) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[clave, cliente.nombre.trim(), cliente.telefono || '', cliente.email || '', cliente.direccion || '']]);
        return 'Cliente actualizado correctamente';
      }
    }
    return 'Cliente no encontrado';
  } catch (err) { return 'Error: ' + err.message; }
}

function obtenerClientes() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(HOJA_CLIENTES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const fotoUrlIdx = headers.indexOf('FotoURL');
  return data.slice(1).map(r => ({
    clave: r[0] || '',
    nombre: r[1] || '',
    telefono: r[2] || '',
    email: r[3] || '',
    direccion: r[4] || '',
    estado: r[5] || 'ACTIVO',
    fotoUrl: fotoUrlIdx >= 0 ? (r[fotoUrlIdx] || '') : ''
  })).filter(c => c.clave);
}

function eliminarCliente(clave) {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_CLIENTES);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === clave) {
      sh.getRange(i + 1, 6).setValue('INACTIVO');
      return 'Cliente dado de baja correctamente';
    }
  }
  throw 'Cliente no encontrado';
}

function reactivarCliente(clave) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(HOJA_CLIENTES);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).toUpperCase() === clave.toUpperCase()) {
      sheet.getRange(i + 1, 6).setValue('ACTIVO');
      return 'Cliente reactivado correctamente';
    }
  }
  throw new Error('Cliente no encontrado');
}

/* ------------------- FOTO CLIENTE ------------------- */
function _getOrCreateFolderByName_(parentFolder, name) {
  var folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(name);
}

function _getClientesFotosFolder_() {
  return _getOrCreateFolderByName_(DriveApp.getRootFolder(), 'CLIENTES_FOTOS');
}

function _asegurarColumnasFotoClientes_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var fotoFileIdCol = headers.indexOf('FotoFileId');
  var fotoUrlCol = headers.indexOf('FotoURL');
  if (fotoFileIdCol === -1) {
    var nextCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextCol).setValue('FotoFileId');
    headers.push('FotoFileId');
    fotoFileIdCol = headers.indexOf('FotoFileId');
  }
  if (fotoUrlCol === -1) {
    var nextCol2 = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextCol2).setValue('FotoURL');
    headers.push('FotoURL');
    fotoUrlCol = headers.indexOf('FotoURL');
  }
  return { fotoFileIdCol: fotoFileIdCol, fotoUrlCol: fotoUrlCol };
}

function subirFotoCliente(payload) {
  try {
    var claveCliente = (payload.claveCliente || '').toString().trim().toUpperCase();
    var base64Data = payload.base64;
    var mimeType = payload.mimeType || 'image/jpeg';
    if (!claveCliente || !base64Data) throw new Error('Datos insuficientes');

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(HOJA_CLIENTES);
    if (!sheet) throw new Error('Hoja CLIENTES no encontrada');

    var colInfo = _asegurarColumnasFotoClientes_(sheet);
    var fotoFileIdCol = colInfo.fotoFileIdCol;
    var fotoUrlCol = colInfo.fotoUrlCol;

    var data = sheet.getDataRange().getValues();
    var fila = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).toUpperCase() === claveCliente) {
        fila = i;
        break;
      }
    }
    if (fila === -1) throw new Error('Cliente no encontrado: ' + claveCliente);

    var prevFileId = data[fila][fotoFileIdCol];
    if (prevFileId) {
      try { DriveApp.getFileById(String(prevFileId)).setTrashed(true); } catch (e) {}
    }

    var folder = _getClientesFotosFolder_();
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, claveCliente + '.jpg');
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    var fotoUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000';

    sheet.getRange(fila + 1, fotoFileIdCol + 1).setValue(fileId);
    sheet.getRange(fila + 1, fotoUrlCol + 1).setValue(fotoUrl);

    return { ok: true, fileId: fileId, fotoUrl: fotoUrl };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/* ------------------- PRODUCTOS ------------------- */
function registrarProducto(producto) {
  try {
    if (!producto || !producto.codigo || !producto.nombre) return "Datos incompletos.";
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!sheet) sheet = ss.insertSheet(HOJA_PRODUCTOS);
    if (!sheet.getLastRow()) sheet.getRange(1, 1, 1, 6).setValues([["Código", "Nombre", "Unidad", "Grupo", "Stock Mínimo", "Fecha Creación"]]);
    const datos = sheet.getDataRange().getValues();
    const codigoNormalizado = String(producto.codigo).trim().toUpperCase();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && String(datos[i][0]).trim().toUpperCase() === codigoNormalizado) return "Ya existe un producto con este código.";
    }
    sheet.appendRow([codigoNormalizado, producto.nombre.trim(), producto.unidad || "Unidades", producto.grupo || "General", parseInt(producto.stockMin) || 0, new Date()]);
    return "Producto registrado correctamente.";
  } catch (error) { return `Error: ${error.message}`; }
}

/* ------------------- PRECIOS (NUEVO, SIN AFECTAR OTROS MODULOS) ------------------- */
function _getHeaderIndex_(headers, name) {
  const target = String(name || '').trim().toUpperCase();
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim().toUpperCase() === target) return i;
  }
  return -1;
}

function obtenerPrecioProducto(codigo) {
  try {
    if (!codigo) return 0;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!sh || sh.getLastRow() < 2) return 0;

    const values = sh.getDataRange().getValues();
    const headers = values[0];

    const idxPrecio = _getHeaderIndex_(headers, "Precio");
    const idxCodigo = _getHeaderIndex_(headers, "Código");
    const cIdx = idxCodigo >= 0 ? idxCodigo : 0;

    if (idxPrecio < 0) return 0;

    const cod = String(codigo).trim().toUpperCase();
    for (let i = 1; i < values.length; i++) {
      const rowCod = String(values[i][cIdx] || '').trim().toUpperCase();
      if (rowCod === cod) {
        const p = parseFloat(values[i][idxPrecio]);
        return isNaN(p) ? 0 : p;
      }
    }
    return 0;
  } catch (e) {
    return 0;
  }
}

function obtenerPrecioProductoPorCodigo(codigo) {
  return obtenerPrecioProducto(codigo);
}

function obtenerPreciosProductos(codigos) {
  try {
    if (!Array.isArray(codigos) || codigos.length === 0) return {};
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!sh || sh.getLastRow() < 2) return {};

    const values = sh.getDataRange().getValues();
    const headers = values[0];
    const idxCodigo = _getHeaderIndex_(headers, "Código");
    const idxPrecio = _getHeaderIndex_(headers, "Precio");
    const cIdx = idxCodigo >= 0 ? idxCodigo : 0;
    if (idxPrecio < 0) return {};

    const wanted = new Set(codigos.map(c => String(c).trim().toUpperCase()));
    const map = {};

    for (let i = 1; i < values.length; i++) {
      const cod = String(values[i][cIdx] || '').trim().toUpperCase();
      if (!wanted.has(cod)) continue;
      const p = parseFloat(values[i][idxPrecio]);
      map[cod] = isNaN(p) ? 0 : p;
    }
    return map;
  } catch (e) {
    return {};
  }
}

/* ------------------- KITS ------------------- */
function obtenerCodigoKit() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HOJA_KITS);
  if (!sheet) sheet = ss.insertSheet(HOJA_KITS);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([["CodigoKit","Nombre","Descripcion","Estado","Fecha Creacion"]]);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'KIT001';
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;
  data.forEach(r => {
    const val = r[0];
    if (val && typeof val === 'string' && val.toUpperCase().startsWith('KIT')) {
      const num = parseInt(val.replace(/[^0-9]/g, ''), 10);
      if (!isNaN(num)) maxNum = Math.max(maxNum, num);
    }
  });
  return 'KIT' + String(maxNum + 1).padStart(3, '0');
}

function guardarKit(kit) {
  try {
    if (!kit || !kit.codigo || !kit.nombre) return { ok: false, mensaje: 'Código y nombre son obligatorios' };
    if (!Array.isArray(kit.componentes) || kit.componentes.length === 0) return { ok: false, mensaje: 'Agregue al menos un componente' };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let kitsSheet = ss.getSheetByName(HOJA_KITS);
    let compSheet = ss.getSheetByName(HOJA_KIT_COMPONENTES);
    if (!kitsSheet) kitsSheet = ss.insertSheet(HOJA_KITS);
    if (!compSheet) compSheet = ss.insertSheet(HOJA_KIT_COMPONENTES);
    if (kitsSheet.getLastRow() === 0) kitsSheet.getRange(1, 1, 1, 5).setValues([["CodigoKit","Nombre","Descripcion","Estado","Fecha Creacion"]]);
    if (compSheet.getLastRow() === 0) compSheet.getRange(1, 1, 1, 3).setValues([["CodigoKit","CodigoProducto","Cantidad"]]);

    const codigo = kit.codigo.toUpperCase().trim();
    const nombre = kit.nombre.trim();
    const descripcion = kit.descripcion || '';
    const estado = kit.estado || 'ACTIVO';

    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    const prodData = prodSheet ? prodSheet.getDataRange().getValues().slice(1) : [];
    const prodSet = new Set(prodData.filter(r => r[0]).map(r => r[0].toString().toUpperCase()));

    const compMap = {};
    for (const c of kit.componentes) {
      const codProd = (c.codigo || '').toUpperCase().trim();
      const cant = parseFloat(c.cantidad);
      if (!codProd) return { ok: false, mensaje: 'Componente sin código' };
      if (!(cant > 0)) return { ok: false, mensaje: `Cantidad inválida para ${codProd}` };
      if (!prodSet.has(codProd)) return { ok: false, mensaje: `Producto no existe: ${codProd}` };
      compMap[codProd] = (compMap[codProd] || 0) + cant;
    }
    const comps = Object.keys(compMap).map(c => [codigo, c, compMap[c]]);

    const kitsData = kitsSheet.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < kitsData.length; i++) {
      if (kitsData[i][0] && kitsData[i][0].toString().toUpperCase() === codigo) { row = i + 1; break; }
    }
    if (row === -1) { kitsSheet.appendRow([codigo, nombre, descripcion, estado, new Date()]); }
    else { kitsSheet.getRange(row, 1, 1, 4).setValues([[codigo, nombre, descripcion, estado]]); }

    const compData = compSheet.getDataRange().getValues();
    for (let i = compData.length; i >= 2; i--) {
      if (compData[i - 1][0] && compData[i - 1][0].toString().toUpperCase() === codigo) { compSheet.deleteRow(i); }
    }
    if (comps.length > 0) compSheet.getRange(compSheet.getLastRow() + 1, 1, comps.length, 3).setValues(comps);

    return { ok: true, mensaje: 'Kit guardado correctamente' };
  } catch (e) { return { ok: false, mensaje: e.message }; }
}

function obtenerKits() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const kitsSheet = ss.getSheetByName(HOJA_KITS);
    const compSheet = ss.getSheetByName(HOJA_KIT_COMPONENTES);
    if (!kitsSheet || kitsSheet.getLastRow() < 2) return [];
    const kitsData = kitsSheet.getDataRange().getValues().slice(1);
    const compsData = compSheet ? compSheet.getDataRange().getValues().slice(1) : [];

    const stockList = obtenerStock();
    const stockMap = {};
    stockList.forEach(p => { stockMap[p.codigo.toUpperCase()] = p.cantidad; });

    const compMap = {};
    compsData.forEach(r => {
      const k = String(r[0]).toUpperCase();
      const p = String(r[1]).toUpperCase();
      const c = parseFloat(r[2]) || 0;
      if (!compMap[k]) compMap[k] = [];
      compMap[k].push({ codigo: p, cantidad: c });
    });

    return kitsData.filter(r => r[0]).map(r => {
      const codigo = String(r[0]).toUpperCase();
      const comps = compMap[codigo] || [];
      const { disponible, limitante } = calcularDisponibilidadKit(comps, stockMap);
      return { codigo, nombre: r[1] || '', descripcion: r[2] || '', estado: r[3] || 'ACTIVO', disponible, limitante };
    });
  } catch (e) { return []; }
}

function calcularDisponibilidadKit(componentes, stockMap) {
  if (!componentes || componentes.length === 0) return { disponible: 0, limitante: 'Sin componentes' };
  let minDisp = Infinity;
  let limitante = '';
  componentes.forEach(c => {
    const stock = stockMap[c.codigo.toUpperCase()] || 0;
    const disp = Math.floor(stock / (c.cantidad || 1));
    if (disp < minDisp) { minDisp = disp; limitante = c.codigo; }
  });
  if (!isFinite(minDisp)) minDisp = 0;
  return { disponible: minDisp, limitante };
}

function obtenerKitDetalle(codigoKit) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const kitsSheet = ss.getSheetByName(HOJA_KITS);
    const compSheet = ss.getSheetByName(HOJA_KIT_COMPONENTES);
    if (!kitsSheet) return null;
    const cod = String(codigoKit).toUpperCase();
    const kitsData = kitsSheet.getDataRange().getValues();
    let kitRow = null;
    for (let i = 1; i < kitsData.length; i++) {
      if (kitsData[i][0] && String(kitsData[i][0]).toUpperCase() === cod) { kitRow = kitsData[i]; break; }
    }
    if (!kitRow) return null;

    const compsData = compSheet ? compSheet.getDataRange().getValues().slice(1) : [];
    const comps = compsData.filter(r => String(r[0]).toUpperCase() === cod)
      .map(r => ({ codigo: r[1], cantidad: r[2], nombre: obtenerNombreProducto(r[1]) }));

    const stockList = obtenerStock();
    const stockMap = {};
    stockList.forEach(p => { stockMap[p.codigo.toUpperCase()] = p.cantidad; });
    const { disponible, limitante } = calcularDisponibilidadKit(comps, stockMap);

    return { codigo: kitRow[0], nombre: kitRow[1], descripcion: kitRow[2] || '', estado: kitRow[3] || 'ACTIVO', disponible, limitante, componentes: comps };
  } catch (e) { return null; }
}

function cambiarEstadoKit(codigo, estado) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const kitsSheet = ss.getSheetByName(HOJA_KITS);
    const data = kitsSheet.getDataRange().getValues();
    const cod = String(codigo).toUpperCase();
    for(let i=1; i<data.length; i++) {
      if(data[i][0] && String(data[i][0]).toUpperCase() === cod) {
        kitsSheet.getRange(i+1, 4).setValue(estado);
        return { ok: true };
      }
    }
    return { ok: false };
  } catch(e) { return { ok: false }; }
}

function buscarKitPorCodigo(texto) {
  try {
    if (!texto || texto.trim().length < 1) return [];
    const kits = obtenerKits();
    const t = texto.toString().toUpperCase().trim();
    return kits.filter(k => k.codigo.toUpperCase().startsWith(t)).slice(0, 10);
  } catch (e) { return []; }
}

function exportarKitsCSV() {
  try {
    const kits = obtenerKits();
    if (kits.length === 0) return null;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const compSheet = ss.getSheetByName(HOJA_KIT_COMPONENTES);
    const compData = compSheet ? compSheet.getDataRange().getValues().slice(1) : [];

    let csv = "\uFEFF";
    csv += "Tipo,Codigo,Nombre/Descripcion/Estado/Limitante/Disponible,CodigoProducto,Cantidad\n";
    kits.forEach(k => {
      csv += `KIT,${k.codigo},"${k.nombre}","${k.descripcion || ''}",${k.estado},${k.limitante || ''},${k.disponible || 0}\n`;
      compData.filter(r => (r[0] || '').toString().toUpperCase() === k.codigo.toUpperCase())
        .forEach(r => { csv += `COMP,${k.codigo},,,${r[1]},${r[2]}\n`; });
    });

    const fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
    const nombreArchivo = `Kits_${fechaHora}.csv`;
    const blob = Utilities.newBlob(csv, 'text/csv; charset=utf-8', nombreArchivo);
    const archivo = DriveApp.getRootFolder().createFile(blob);
    return archivo.getUrl();
  } catch (e) { return null; }
}

/* ------------------- CAJA REGISTRADORA (EXTENDIDA A 19 COLS) ------------------- */
function generarIdCaja(tipo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HOJA_CAJA);
  if (!sheet) sheet = ss.insertSheet(HOJA_CAJA);

  const needCols = 19;
  const lastCol = sheet.getLastColumn();
  if (lastCol < needCols) sheet.insertColumnsAfter(lastCol, needCols - lastCol);

  const headers = [[
    "IDREGISTRO","Fecha Captura","Cliente","Agente","Obra","Remisión","Tipo",
    "Producto","Cantidad","Fecha Renta Inicio","Fecha Renta Fin","Usuario","Timestamp","KitCodigo",
    "PrecioUnit","DescPct","DescMonto","SubTotalLinea","ImporteLinea"
  ]];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, needCols).setValues(headers);
  } else {
    // Mantener header consistente sin afectar otras filas
    sheet.getRange(1, 1, 1, needCols).setValues(headers);
  }

  const prefijo = tipo === TIPOS_MOVIMIENTO.SALIDA_RENTA
    ? 'RENT'
    : tipo === TIPOS_MOVIMIENTO.SALIDA_VENTA
      ? 'VENT'
      : 'CJ';

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return prefijo + '01';

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;
  data.forEach(row => {
    const val = row[0];
    if (val && typeof val === 'string' && val.toUpperCase().startsWith(prefijo)) {
      const num = parseInt(val.replace(/[^0-9]/g, ''), 10);
      if (!isNaN(num)) maxNum = Math.max(maxNum, num);
    }
  });

  const nextNum = maxNum + 1;
  return prefijo + String(nextNum).padStart(2, '0');
}

function registrarCaja(payload) {
  let lock;
  try {
    if (!payload || !payload.cabecera || !Array.isArray(payload.productos)) {
      return { ok: false, errores: ['Payload inválido'] };
    }
    const { cabecera, productos } = payload;
    const errores = [];

    if (!cabecera.cliente) errores.push('Cliente es obligatorio');
    if (!cabecera.tipo) errores.push('Tipo es obligatorio');
    if (!productos.length) errores.push('Debe agregar al menos un producto');
    if (errores.length) return { ok: false, errores };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    const cliSheet = ss.getSheetByName(HOJA_CLIENTES);
    const cajaSheet = ss.getSheetByName(HOJA_CAJA);

    if (!prodSheet || !cliSheet || !cajaSheet) return { ok: false, errores: ['Faltan hojas requeridas'] };

    // asegurar 19 columnas en caja
    const needCols = 19;
    const lastCol = cajaSheet.getLastColumn();
    if (lastCol < needCols) cajaSheet.insertColumnsAfter(lastCol, needCols - lastCol);

    const clienteActivo = obtenerClientes().some(c => String(c.clave).toUpperCase() === String(cabecera.cliente).toUpperCase() && String(c.estado).toUpperCase() === 'ACTIVO');
    if (!clienteActivo) errores.push('El cliente no existe o está INACTIVO');

    const productosMap = new Set(prodSheet.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => String(r[0]).toUpperCase()));
    const movimientos = [];
    const hoy = new Date();
    const idRegistro = generarIdCaja(cabecera.tipo);

    productos.forEach((p, idx) => {
      const linea = idx + 1;
      const codigo = String(p.producto || '').toUpperCase();
      const cantidad = parseFloat(p.cantidad);
      if (!codigo) errores.push(`Línea ${linea}: producto requerido`);
      if (!(cantidad > 0)) errores.push(`Línea ${linea}: cantidad inválida`);
      if (!productosMap.has(codigo)) errores.push(`Línea ${linea}: producto no encontrado (${codigo})`);
      if (cabecera.tipo === TIPOS_MOVIMIENTO.SALIDA_RENTA) {
        if (!p.fechaInicio || !p.fechaFin) errores.push(`Línea ${linea}: fechas requeridas para renta`);
        if (p.fechaInicio && p.fechaFin && new Date(p.fechaInicio) > new Date(p.fechaFin)) errores.push(`Línea ${linea}: fecha inicio > fecha fin`);
      }
      movimientos.push({
        codigo,
        fecha: Utilities.formatDate(hoy, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        tipo: cabecera.tipo,
        cantidad,
        observaciones: `Caja ${idRegistro} ${p.kitCodigo ? '(Kit ' + p.kitCodigo + ')' : ''} ${cabecera.remision || ''}`.trim()
      });
    });

    if (errores.length) return { ok: false, errores };

    lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, errores: ['Sistema ocupado, intente nuevamente.'] };

    const resMov = registrarMovimientosBatch(movimientos);
    if (!resMov.ok) return resMov;

    const user = Session.getActiveUser().getEmail() || 'Sistema';
    const filasCaja = productos.map(p => ([
      idRegistro,
      hoy,
      cabecera.cliente,
      cabecera.agente || '',
      cabecera.obra || '',
      cabecera.remision || '',
      cabecera.tipo,
      String(p.producto || '').toUpperCase(),
      parseFloat(p.cantidad) || 0,
      p.fechaInicio || '',
      p.fechaFin || '',
      user,
      new Date(),
      p.kitCodigo || '',
      parseFloat(p.precioUnitario) || 0,
      parseFloat(p.descPct) || 0,
      parseFloat(p.descMonto) || 0,
      parseFloat(p.subTotalLinea) || 0,
      parseFloat(p.importeLinea) || 0
    ]));

    const startRow = cajaSheet.getLastRow() + 1;
    cajaSheet.insertRowsAfter(cajaSheet.getLastRow() || 1, filasCaja.length);
    cajaSheet.getRange(startRow, 1, filasCaja.length, filasCaja[0].length).setValues(filasCaja);

    return { ok: true, idRegistro };
  } catch (error) {
    console.error("Error en registrarCaja:", error);
    return { ok: false, errores: [error.message] };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/* ------------------- MOVIMIENTOS ------------------- */
function registrarMovimientosBatch(movs) {
  let lock;
  try {
    if (!Array.isArray(movs) || movs.length === 0) return { ok: false, errores: ['No hay movimientos'] };
    lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, errores: ['Sistema ocupado'] };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    if (!movSheet) throw new Error("Falta hoja Movimientos");
    if (!movSheet.getLastRow()) {
      movSheet.getRange(1, 1, 1, 8).setValues([["Código", "Fecha", "Tipo", "Cantidad", "Usuario", "Timestamp", "Observaciones", "Stock Resultante"]]);
    }
    const filasParaInsertar = [];
    movs.forEach(m => {
      filasParaInsertar.push([
        String(m.codigo).toUpperCase(),
        new Date(m.fecha + 'T12:00:00'),
        m.tipo,
        m.cantidad,
        Session.getActiveUser().getEmail() || "Sistema",
        new Date(),
        m.observaciones || "",
        null
      ]);
    });
    const startRow = movSheet.getLastRow() + 1;
    movSheet.insertRowsAfter(movSheet.getLastRow() || 1, filasParaInsertar.length);
    movSheet.getRange(startRow, 1, filasParaInsertar.length, filasParaInsertar[0].length).setValues(filasParaInsertar);
    return { ok: true, errores: [] };
  } catch (error) {
    return { ok: false, errores: [error.message] };
  } finally {
    if (lock) try { lock.releaseLock(); } catch (e) {}
  }
}

/* ------------------- CHECK DE RENTAS (CORE) ------------------- */
function obtenerRentaPorId(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const caja = ss.getSheetByName(HOJA_CAJA);

    if (!caja || caja.getLastRow() < 2) throw new Error('No hay registros en Caja');

    if (caja.getLastColumn() < 14) {
      caja.insertColumnAfter(13);
      caja.getRange(1, 14).setValue('KitCodigo');
    }

    const prodMap = obtenerMapaProductosPorCodigo();
    const data = caja.getDataRange().getValues().slice(1);
    const idUpper = String(id).trim().toUpperCase();

    const filas = data.filter(r => {
      const rid = String(r[0] || '').trim().toUpperCase();
      const tipo = normalizarTipoMovimiento(r[6] || '');
      return esTipoRentaFlexible(tipo) && (rid === idUpper);
    });

    if (!filas.length) return null;

    const lineasExpandidas = [];
    filas.forEach(f => {
      const kitCodigo = String(f[13] || '').trim();
      const codigo = String(f[7] || '').toUpperCase();
      const cantidad = parseFloat(f[8]) || 0;
      const nombreProd = prodMap[codigo] || obtenerNombreProducto(codigo);

      lineasExpandidas.push({
        codigo: codigo,
        nombre: nombreProd,
        esperado: cantidad,
        kitCodigo: kitCodigo
      });
    });

    const agrupadas = {};
    lineasExpandidas.forEach(l => {
      const k = l.codigo.toUpperCase();
      if (!agrupadas[k]) {
        agrupadas[k] = {
          codigo: k,
          esperado: 0,
          kitCodigo: l.kitCodigo || '',
          nombre: l.nombre || ''
        };
      }
      agrupadas[k].esperado += l.esperado;
      if (!agrupadas[k].kitCodigo && l.kitCodigo) agrupadas[k].kitCodigo = l.kitCodigo;
    });

    const lineasFinal = Object.values(agrupadas).map(l => ({
      codigo: l.codigo,
      cantidad: Number(l.esperado.toFixed(2)),
      kitCodigo: l.kitCodigo,
      nombre: l.nombre || prodMap[l.codigo] || obtenerNombreProducto(l.codigo)
    }));

    const cab = filas[0];
    const mapaClientes = obtenerMapaClientesPorClave();
    const clienteClave = String(cab[2] || '').toUpperCase();
    const clienteNombre = mapaClientes[clienteClave] || cab[2] || '';

    return {
      id: cab[0],
      fechaCaptura: cab[1] ? formatearFecha(cab[1]) : '',
      cliente: clienteNombre,
      clienteClave,
      agente: cab[3],
      obra: cab[4],
      remision: cab[5],
      tipo: cab[6],
      lineas: lineasFinal
    };

  } catch (e) {
    console.error('obtenerRentaPorId Error:', e);
    throw e.message;
  }
}

function guardarCheckRenta(payload) {
  let lock;
  try {
    if (!payload || !payload.id || !Array.isArray(payload.lineas)) return { ok: false, errores: ['Payload inválido'] };
    const id = String(payload.id).toUpperCase();
    const lineas = payload.lineas;
    lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, errores: ['Sistema ocupado'] };

    const movs = [];
    const hoy = new Date();
    lineas.forEach(l => {
      const bueno = parseFloat(l.bueno) || 0;
      const danado = parseFloat(l.danado) || 0;
      const faltante = parseFloat(l.faltante) || 0;
      const obsBase = `Check renta ID ${id} ${l.kitCodigo ? '(Kit ' + l.kitCodigo + ')' : ''}`.trim();
      const comentarioUser = l.comentario ? ` - ${l.comentario}` : '';
      if (bueno > 0) {
        movs.push({
          codigo: l.codigo.toUpperCase(),
          fecha: Utilities.formatDate(hoy, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          tipo: TIPOS_MOVIMIENTO.INGRESO,
          cantidad: bueno,
          observaciones: `${obsBase} (Buen estado)${comentarioUser}`
        });
      }
      if (danado > 0) {
        movs.push({
          codigo: l.codigo.toUpperCase(),
          fecha: Utilities.formatDate(hoy, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          tipo: TIPOS_MOVIMIENTO.ENTRADA_MANTTO,
          cantidad: danado,
          observaciones: `${obsBase} (Dañado)${comentarioUser}`
        });
      }
      if (faltante > 0) {
        movs.push({
          codigo: l.codigo.toUpperCase(),
          fecha: Utilities.formatDate(hoy, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          tipo: TIPOS_MOVIMIENTO.AJUSTE_NEGATIVO,
          cantidad: faltante,
          observaciones: `${obsBase} (Faltante/Cobrado)${comentarioUser}`
        });
      }
    });

    if (movs.length === 0) return { ok: false, errores: ['No hay cantidades a procesar.'] };
    const resBatch = registrarMovimientosBatch(movs);
    if (!resBatch.ok) return resBatch;
    return { ok: true };
  } catch (e) {
    console.error('guardarCheckRenta:', e);
    return { ok: false, errores: [e.message] };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e) {} }
  }
}

function listarRentasPorEstado(estado) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const caja = ss.getSheetByName(HOJA_CAJA);
    if (!caja || caja.getLastRow() < 2) return [];

    const prodMap = obtenerMapaProductosPorCodigo();
    const data = caja.getDataRange().getValues().slice(1);
    const mapaClientes = obtenerMapaClientesPorClave();
    const rentas = new Map();

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    const hoyMs = hoy.getTime();

    data.forEach(r => {
      const id = String(r[0] || '').trim().toUpperCase();
      const tipo = normalizarTipoMovimiento(r[6] || '');
      if (!esTipoRentaFlexible(tipo)) return;

      const clienteClave = String(r[2] || '').toUpperCase();
      const clienteNombre = mapaClientes[clienteClave] || r[2] || '';

      if (!rentas.has(id)) {
        rentas.set(id, {
          id,
          cliente: clienteNombre,
          clienteClave,
          agente: r[3],
          lines: [],
          maxFechaFinMs: 0
        });
      }

      const rentaActual = rentas.get(id);

      let rawFechaFin = r[10];
      let msFechaFin = 0;

      if (rawFechaFin instanceof Date) {
        let f = new Date(rawFechaFin);
        f.setHours(0, 0, 0, 0);
        msFechaFin = f.getTime();
      } else if (rawFechaFin) {
        let f = new Date(rawFechaFin);
        if(!isNaN(f.getTime())) {
          f.setHours(0, 0, 0, 0);
          msFechaFin = f.getTime();
        }
      }

      if (msFechaFin > rentaActual.maxFechaFinMs) {
        rentaActual.maxFechaFinMs = msFechaFin;
      }

      const codigoProd = String(r[7]).toUpperCase();
      const fechaIni = r[9] ? formatearFecha(r[9]) : '';
      const fechaFinStr = r[10] ? formatearFecha(r[10]) : '';

      rentaActual.lines.push({
        codigo: codigoProd,
        nombre: prodMap[codigoProd] || obtenerNombreProducto(codigoProd),
        cantidad: parseFloat(r[8]) || 0,
        fechaInicio: fechaIni,
        fechaFin: fechaFinStr
      });
    });

    const result = [];

    for (const [id, renta] of rentas) {
      const estadoCalc = calcularEstadoRenta(id, renta.lines);
      renta.estado = estadoCalc;

      renta.esVencida = false;
      renta.venceHoy = false;

      if (estadoCalc !== 'COMPLETA' && renta.maxFechaFinMs > 0) {
        if (hoyMs > renta.maxFechaFinMs) renta.esVencida = true;
        else if (hoyMs === renta.maxFechaFinMs) renta.venceHoy = true;
      }

      delete renta.maxFechaFinMs;

      let incluir = false;
      if (!estado) incluir = true;
      else if (estado === 'COMPLETA') { if (estadoCalc === 'COMPLETA') incluir = true; }
      else if (estado === 'VENCIDA') { if (renta.esVencida) incluir = true; }
      else if (estado === 'VENCE_HOY') { if (renta.venceHoy) incluir = true; }
      else if (estado === 'ACTIVO') { if (estadoCalc === 'ACTIVO' && !renta.esVencida && !renta.venceHoy) incluir = true; }
      else if (estado === 'PENDIENTE') { if (estadoCalc === 'ACTIVO') incluir = true; }

      if (incluir) result.push(renta);
    }

    return result;
  } catch (e) { return []; }
}

function calcularEstadoRenta(id, lineasCaja) {
  const movs = movimientosPorId(id);
  const sumMapa = {};
  movs.forEach(m => {
    if (!sumMapa[m.codigo]) sumMapa[m.codigo] = { procesado: 0 };
    if (m.tipo === TIPOS_MOVIMIENTO.INGRESO || m.tipo === TIPOS_MOVIMIENTO.ENTRADA_MANTTO || m.tipo === TIPOS_MOVIMIENTO.AJUSTE_NEGATIVO || m.tipo === TIPOS_MOVIMIENTO.ENTRADA_RENTA) {
      sumMapa[m.codigo].procesado += m.cantidad;
    }
  });
  const resumenCaja = {};
  lineasCaja.forEach(l => {
    const k = String(l.codigo).toUpperCase();
    resumenCaja[k] = (resumenCaja[k] || 0) + (parseFloat(l.cantidad) || 0);
  });
  for (const codigo in resumenCaja) {
    const esperado = resumenCaja[codigo];
    const procesado = sumMapa[codigo] ? sumMapa[codigo].procesado : 0;
    if (procesado < esperado - 0.01) return 'ACTIVO';
  }
  return 'COMPLETA';
}

function movimientosPorId(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    if (!movSheet || movSheet.getLastRow() < 2) return [];
    const data = movSheet.getDataRange().getValues();
    const idUpper = String(id).toUpperCase();
    const encontrados = [];
    for(let i=1; i<data.length; i++){
      const obs = String(data[i][6] || '').toUpperCase();
      if(obs.includes(idUpper)) {
        encontrados.push({
          codigo: String(data[i][0]).toUpperCase(),
          fecha: data[i][1],
          tipo: normalizarTipoMovimiento(data[i][2]),
          cantidad: parseFloat(data[i][3]) || 0,
          obs: obs
        });
      }
    }
    return encontrados;
  } catch (e) { return []; }
}

function debugValidarRentaToSheet(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dbgName = 'DEBUG_RENTAS';
  let dbg = ss.getSheetByName(dbgName);
  if (!dbg) dbg = ss.insertSheet(dbgName);
  dbg.clear();

  dbg.appendRow(['ID BUSCADO', id ?? '']);
  dbg.getRange(1, 1, 1, 2).setBackground("#FFA500").setFontWeight("bold");

  try {
    const idStr = String(id ?? '').trim();
    if (!idStr) {
      dbg.appendRow(['Error', 'ID vacío o indefinido']);
      return 'ID vacío';
    }

    const caja = ss.getSheetByName(HOJA_CAJA);
    if (!caja) throw new Error('Hoja CAJA_REGISTRADORA no encontrada');

    const data = caja.getDataRange().getValues();
    const idUpper = idStr.toUpperCase();

    const rowsFound = [];
    for(let i=1; i<data.length; i++) {
      const rowVal = String(data[i][0] || '').trim().toUpperCase();
      if(rowVal.includes(idUpper)) {
        rowsFound.push({rowIndex: i+1, data: data[i]});
      }
    }

    dbg.appendRow(['Filas encontradas (coincidencia parcial)', rowsFound.length]);
    if (rowsFound.length === 0) {
      dbg.appendRow(['Mensaje', 'No se encontró el ID en CAJA_REGISTRADORA']);
      return 'No rows';
    }

    dbg.appendRow(['']);
    dbg.appendRow(['# ROW', 'ID', 'Fecha', 'Cliente', 'Tipo RAW', 'Tipo NORM', 'EsRenta?', 'Producto', 'Cantidad', 'KitCodigo']);
    dbg.getRange(dbg.getLastRow(), 1, 1, 10).setBackground("#DDDDDD").setFontWeight("bold");

    rowsFound.forEach(item => {
      const r = item.data;
      const tipoRaw = String(r[6] ?? '');
      const tipoNorm = normalizarTipoMovimiento(tipoRaw);
      dbg.appendRow([
        item.rowIndex,
        r[0] ?? '',
        r[1] ?? '',
        r[2] ?? '',
        tipoRaw,
        tipoNorm,
        esTipoRentaFlexible(tipoRaw),
        r[7] ?? '',
        r[8] ?? '',
        r[13] ?? ''
      ]);
    });

    dbg.autoResizeColumns(1, 10);
    return 'OK';
  } catch (e) {
    dbg.appendRow(['Error Critico', e.message]);
    return 'Error';
  }
}

/* ------------------- CHECK DE VENTAS (CORE) ------------------- */
const HOJA_CHECKS_VENTAS = "CHECKS_VENTAS";
const HOJA_DEPOSITOS = "DEPOSITOS";

function esTipoVentaFlexible(tipo) {
  const t = normalizarTipoMovimiento(tipo);
  return t === TIPOS_MOVIMIENTO.SALIDA_VENTA;
}

function obtenerVentaPorId(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const caja = ss.getSheetByName(HOJA_CAJA);
    if (!caja || caja.getLastRow() < 2) throw new Error('No hay registros en Caja');

    const prodMap = obtenerMapaProductosPorCodigo();
    const data = caja.getDataRange().getValues().slice(1);
    const idUpper = String(id).trim().toUpperCase();

    const filas = data.filter(r => {
      const rid = String(r[0] || '').trim().toUpperCase();
      const tipo = normalizarTipoMovimiento(r[6] || '');
      return esTipoVentaFlexible(tipo) && (rid === idUpper);
    });

    if (!filas.length) return null;

    const lineas = filas.map(f => {
      const codigo = String(f[7] || '').toUpperCase();
      const cantidad = parseFloat(f[8]) || 0;
      const precioUnit = parseFloat(f[14]) || 0;
      const importeLinea = parseFloat(f[18]) || 0;
      return {
        codigo,
        nombre: prodMap[codigo] || obtenerNombreProducto(codigo),
        cantidad,
        precioUnit,
        importeLinea,
        kitCodigo: String(f[13] || '').trim()
      };
    });

    const totalImporte = lineas.reduce((s, l) => s + l.importeLinea, 0);
    const cab = filas[0];
    const mapaClientes = obtenerMapaClientesPorClave();
    const clienteClave = String(cab[2] || '').toUpperCase();
    const clienteNombre = mapaClientes[clienteClave] || cab[2] || '';
    const estadoPago = calcularEstadoVenta(idUpper);
    const depositos = obtenerDepositosPorVenta(idUpper);
    const totalDepositos = Number(depositos.reduce((s, d) => s + d.monto, 0).toFixed(2));
    const totalImporteRedondeado = Number(totalImporte.toFixed(2));

    return {
      id: cab[0],
      fechaCaptura: cab[1] ? formatearFecha(cab[1]) : '',
      cliente: clienteNombre,
      clienteClave,
      agente: cab[3],
      obra: cab[4],
      remision: cab[5],
      tipo: cab[6],
      lineas,
      totalImporte: totalImporteRedondeado,
      totalDepositos,
      saldo: _calcularSaldo_(totalImporteRedondeado, totalDepositos),
      depositos,
      estadoPago
    };
  } catch (e) {
    console.error('obtenerVentaPorId Error:', e);
    throw e.message;
  }
}

function calcularEstadoVenta(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const checksSheet = ss.getSheetByName(HOJA_CHECKS_VENTAS);
    if (!checksSheet || checksSheet.getLastRow() < 2) return 'CON_SALDO';
    const idUpper = String(id).trim().toUpperCase();
    const data = checksSheet.getDataRange().getValues().slice(1);
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toUpperCase() === idUpper) return 'SIN_SALDO';
    }
    return 'CON_SALDO';
  } catch (e) { return 'CON_SALDO'; }
}

function guardarCheckVenta(payload) {
  let lock;
  try {
    if (!payload || !payload.id) return { ok: false, errores: ['Payload inválido'] };
    const id = String(payload.id).toUpperCase();
    lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, errores: ['Sistema ocupado'] };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let checksSheet = ss.getSheetByName(HOJA_CHECKS_VENTAS);
    if (!checksSheet) {
      checksSheet = ss.insertSheet(HOJA_CHECKS_VENTAS);
      checksSheet.getRange(1, 1, 1, 4).setValues([['IDREGISTRO', 'Fecha', 'Usuario', 'Comentario']]);
    }

    if (checksSheet.getLastRow() >= 2) {
      const existing = checksSheet.getDataRange().getValues().slice(1);
      for (let i = 0; i < existing.length; i++) {
        if (String(existing[i][0] || '').trim().toUpperCase() === id) {
          return { ok: false, errores: ['Esta venta ya fue marcada como SIN SALDO.'] };
        }
      }
    }

    const hoy = new Date();
    const user = Session.getActiveUser().getEmail() || 'Sistema';
    checksSheet.appendRow([id, hoy, user, payload.comentario || '']);
    return { ok: true };
  } catch (e) {
    console.error('guardarCheckVenta:', e);
    return { ok: false, errores: [e.message] };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/* ------------------- DEPÓSITOS DE VENTAS ------------------- */

function _asegurarHojaDepositos_(ss) {
  let sheet = ss.getSheetByName(HOJA_DEPOSITOS);
  if (!sheet) {
    sheet = ss.insertSheet(HOJA_DEPOSITOS);
    sheet.getRange(1, 1, 1, 6).setValues([['VENTA_ID', 'Fecha', 'Monto', 'Tipo', 'Comentario', 'Usuario']]);
  }
  return sheet;
}

function registrarDeposito(payload) {
  let lock;
  try {
    if (!payload || !payload.ventaId || !payload.monto || !payload.tipo) {
      return { ok: false, error: 'Datos incompletos' };
    }
    const monto = parseFloat(payload.monto);
    if (isNaN(monto) || monto <= 0) return { ok: false, error: 'Monto inválido' };

    lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: 'Sistema ocupado' };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = _asegurarHojaDepositos_(ss);
    const user = Session.getActiveUser().getEmail() || 'Sistema';
    sheet.appendRow([
      String(payload.ventaId).toUpperCase(),
      new Date(),
      monto,
      String(payload.tipo),
      String(payload.comentario || ''),
      user
    ]);
    return { ok: true };
  } catch (e) {
    console.error('registrarDeposito:', e);
    return { ok: false, error: e.message };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e) {} }
  }
}

function obtenerDepositosPorVenta(ventaId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_DEPOSITOS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const idUpper = String(ventaId).trim().toUpperCase();
    const data = sheet.getDataRange().getValues().slice(1);
    return data
      .filter(r => String(r[0] || '').trim().toUpperCase() === idUpper)
      .map(r => ({
        ventaId: r[0],
        fecha: r[1] ? formatearFecha(r[1]) : '',
        monto: parseFloat(r[2]) || 0,
        tipo: String(r[3] || ''),
        comentario: String(r[4] || ''),
        usuario: String(r[5] || '')
      }));
  } catch (e) { return []; }
}

function _calcularSaldo_(totalImporte, totalDepositos) {
  return Number(Math.max(0, totalImporte - totalDepositos).toFixed(2));
}

function _calcularTotalDepositos_(ss, ventaId) {
  try {
    const sheet = ss.getSheetByName(HOJA_DEPOSITOS);
    if (!sheet || sheet.getLastRow() < 2) return 0;
    const idUpper = String(ventaId).trim().toUpperCase();
    const data = sheet.getDataRange().getValues().slice(1);
    return data
      .filter(r => String(r[0] || '').trim().toUpperCase() === idUpper)
      .reduce((sum, r) => sum + (parseFloat(r[2]) || 0), 0);
  } catch (e) { return 0; }
}

function listarVentasPorEstado(estado) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const caja = ss.getSheetByName(HOJA_CAJA);
    if (!caja || caja.getLastRow() < 2) return [];

    const prodMap = obtenerMapaProductosPorCodigo();
    const data = caja.getDataRange().getValues().slice(1);
    const mapaClientes = obtenerMapaClientesPorClave();
    const ventas = new Map();

    data.forEach(r => {
      const id = String(r[0] || '').trim().toUpperCase();
      const tipo = normalizarTipoMovimiento(r[6] || '');
      if (!esTipoVentaFlexible(tipo)) return;

      const clienteClave = String(r[2] || '').toUpperCase();
      const clienteNombre = mapaClientes[clienteClave] || r[2] || '';

      if (!ventas.has(id)) {
        ventas.set(id, {
          id,
          cliente: clienteNombre,
          clienteClave,
          agente: r[3],
          fechaCaptura: r[1] ? formatearFecha(r[1]) : '',
          lines: [],
          totalImporte: 0
        });
      }

      const ventaActual = ventas.get(id);
      const codigoProd = String(r[7]).toUpperCase();
      const importeLinea = parseFloat(r[18]) || 0;
      ventaActual.totalImporte += importeLinea;
      ventaActual.lines.push({
        codigo: codigoProd,
        nombre: prodMap[codigoProd] || obtenerNombreProducto(codigoProd),
        cantidad: parseFloat(r[8]) || 0,
        importeLinea
      });
    });

    const liquidadas = new Set();
    const checksSheet = ss.getSheetByName(HOJA_CHECKS_VENTAS);
    if (checksSheet && checksSheet.getLastRow() >= 2) {
      checksSheet.getDataRange().getValues().slice(1).forEach(r => {
        const idCheck = String(r[0] || '').trim().toUpperCase();
        if (idCheck) liquidadas.add(idCheck);
      });
    }

    const result = [];
    for (const [id, venta] of ventas) {
      venta.totalImporte = Number(venta.totalImporte.toFixed(2));
      venta.estadoPago = liquidadas.has(id) ? 'SIN_SALDO' : 'CON_SALDO';
      const totalDepositos = Number(_calcularTotalDepositos_(ss, id).toFixed(2));
      venta.totalDepositos = totalDepositos;
      venta.saldo = _calcularSaldo_(venta.totalImporte, totalDepositos);
      let incluir = false;
      if (!estado || estado === 'TODAS') incluir = true;
      else if (estado === 'CON_SALDO') incluir = (venta.estadoPago === 'CON_SALDO');
      else if (estado === 'SIN_SALDO') incluir = (venta.estadoPago === 'SIN_SALDO');
      if (incluir) result.push(venta);
    }
    return result;
  } catch (e) { return []; }
}

/* ------------------- PARAMETROS (CON LOGO) ------------------- */
function _asegurarHojaParametros_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HOJA_PARAMETROS);

  const headers = [[
    "Nombre", "RFC", "Dirección", "Web", "Colonia", "Ciudad", "Estado", "Pais", "CP", "Regimen", "Email", "Teléfono",
    "LogoFileId", "LogoUrl"
  ]];

  if (!sheet) {
    sheet = ss.insertSheet(HOJA_PARAMETROS);
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  } else {
    const needCols = headers[0].length;
    const lastCol = sheet.getLastColumn();
    if (lastCol < needCols) sheet.insertColumnsAfter(lastCol, needCols - lastCol);
    sheet.getRange(1, 1, 1, needCols).setValues(headers);
  }

  if (sheet.getLastRow() < 2) sheet.appendRow(Array(14).fill(""));
  return sheet;
}

function obtenerParametros() {
  try {
    const sheet = _asegurarHojaParametros_();
    const data = sheet.getRange(2, 1, 1, 14).getValues()[0];
    return {
      nombre: data[0] || "",
      rfc: data[1] || "",
      direccion: data[2] || "",
      web: data[3] || "",
      colonia: data[4] || "",
      ciudad: data[5] || "",
      estado: data[6] || "",
      pais: data[7] || "",
      cp: data[8] || "",
      regimen: data[9] || "",
      email: data[10] || "",
      telefono: data[11] || "",
      logoFileId: data[12] || "",
      logoUrl: data[13] || ""
    };
  } catch (e) {
    return { error: e.message };
  }
}

function guardarParametros(params) {
  try {
    const sheet = _asegurarHojaParametros_();
    const current = sheet.getRange(2, 1, 1, 14).getValues()[0];

    const logoFileId = current[12] || "";
    const logoUrl = current[13] || "";

    const rowData = [
      params.nombre || "", params.rfc || "", params.direccion || "", params.web || "",
      params.colonia || "", params.ciudad || "", params.estado || "", params.pais || "",
      params.cp || "", params.regimen || "", params.email || "", params.telefono || "",
      logoFileId, logoUrl
    ];

    sheet.getRange(2, 1, 1, 14).setValues([rowData]);
    return { ok: true, mensaje: "Parámetros actualizados correctamente." };
  } catch (e) {
    return { ok: false, mensaje: "Error: " + e.message };
  }
}

function subirLogoEmpresa(payload) {
  try {
    if (!payload || !payload.base64 || !payload.mimeType) {
      return { ok: false, mensaje: "Payload inválido" };
    }

    const sheet = _asegurarHojaParametros_();

    const b64 = String(payload.base64).split(",").pop();
    const bytes = Utilities.base64Decode(b64);

    const mt = String(payload.mimeType).toLowerCase();
    if (mt !== "image/png" && mt !== "image/jpeg") {
      return { ok: false, mensaje: "Tipo inválido. Solo PNG o JPG." };
    }

    const ext = mt === "image/png" ? "png" : "jpg";
    const fileName = `LOGO_EMPRESA_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss")}.${ext}`;
    const blob = Utilities.newBlob(bytes, mt, fileName);

    const folderName = "INVENTARIO_LOGOS";
    const it = DriveApp.getFoldersByName(folderName);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);

    const oldId = sheet.getRange(2, 13).getValue();
    if (oldId) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch(e) {} }

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    const logoUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w1000`;

    sheet.getRange(2, 13).setValue(fileId);
    sheet.getRange(2, 14).setValue(logoUrl);

    return { ok: true, fileId, logoUrl };
  } catch (e) {
    return { ok: false, mensaje: "Error al subir: " + e.message };
  }
}

/* ------------------- HELPERS ------------------- */
function calcularStock(codigoProducto) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    if (!movSheet) return 0;
    const codigo = String(codigoProducto).trim().toUpperCase();
    const movimientos = movSheet.getDataRange().getValues();
    let stock = 0;
    for (let i = 1; i < movimientos.length; i++) {
      const mov = movimientos[i];
      if (!mov[0] || String(mov[0]).trim().toUpperCase() !== codigo) continue;
      const tipo = mov[2] ? String(mov[2]).toUpperCase() : "";
      const cantidad = parseFloat(mov[3]) || 0;
      if (esEntrada(tipo)) stock += cantidad;
      else if (esSalida(tipo)) stock -= cantidad;
      else if (tipo === TIPOS_MOVIMIENTO.AJUSTE) stock += cantidad;
    }
    return Math.round(stock * 100) / 100;
  } catch (error) { return 0; }
}

function obtenerStock() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!prodSheet) return [];
    const productos = prodSheet.getDataRange().getValues();
    if (productos.length <= 1) return [];
    const stock = [];
    for (let i = 1; i < productos.length; i++) {
      const [codigo, nombre, unidad, grupo, stockMin] = productos[i];
      if (codigo && nombre) {
        stock.push({
          codigo: String(codigo),
          nombre: String(nombre),
          unidad: unidad || "Unidades",
          grupo: grupo || "General",
          stockMin: Math.max(0, parseInt(stockMin) || 0),
          cantidad: calcularStock(codigo)
        });
      }
    }
    return stock.sort((a, b) => a.nombre.localeCompare(b.nombre));
  } catch (error) { return []; }
}

function buscarProductoPorCodigo(codigo) {
  try {
    if (!codigo || String(codigo).trim().length < 1) return [];
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!sheet) return [];
    const datos = sheet.getDataRange().getValues();
    const textoBusqueda = String(codigo).toUpperCase().trim();
    const encontrados = [];
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (fila[0] && String(fila[0]).toUpperCase().startsWith(textoBusqueda)) {
        encontrados.push({ codigo: fila[0], nombre: fila[1], unidad: fila[2] || "Unidades", grupo: fila[3] || "General" });
      }
    }
    return encontrados.slice(0, 10);
  } catch (error) { return []; }
}

function obtenerHistorial(filtros) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!movSheet || !prodSheet) throw new Error("Las hojas del sistema no existen.");
    const movimientos = movSheet.getDataRange().getValues();
    const productos = prodSheet.getDataRange().getValues();
    const prodMap = {};
    for (let i = 1; i < productos.length; i++) {
      if (productos[i][0]) prodMap[productos[i][0].toString().toUpperCase()] = productos[i][1];
    }
    const fechaDesde = new Date(filtros.fechaDesde + 'T00:00:00');
    const fechaHasta = new Date(filtros.fechaHasta + 'T23:59:59');
    const resultado = [];
    for (let i = 1; i < movimientos.length; i++) {
      const mov = movimientos[i];
      if (!mov[0] || !mov[1]) continue;
      try {
        const fechaMov = new Date(mov[1]);
        const tipoMov = mov[2] ? mov[2].toString().toUpperCase() : "";
        if (fechaMov >= fechaDesde && fechaMov <= fechaHasta) {
          if (!filtros.tipo || tipoMov === filtros.tipo.toUpperCase()) {
            const codigoProducto = mov[0].toString().toUpperCase();
            resultado.push({
              codigo: mov[0],
              fecha: formatearFecha(fechaMov),
              tipo: tipoMov,
              tipoTexto: tipoMov,
              cantidad: parseFloat(mov[3]) || 0,
              producto: prodMap[codigoProducto] || "Producto no encontrado",
              observaciones: mov[6] || "",
              usuario: mov[4] || "N/A"
            });
          }
        }
      } catch (dateError) { continue; }
    }
    return resultado.sort((a, b) => {
      const fechaA = new Date(a.fecha.split('/').reverse().join('-'));
      const fechaB = new Date(b.fecha.split('/').reverse().join('-'));
      return fechaB - fechaA;
    });
  } catch (error) { return []; }
}

function obtenerNombreProducto(c) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!prodSheet) return "";
    const data = prodSheet.getDataRange().getValues();
    const cod = String(c).toUpperCase();
    for(let i=1; i<data.length; i++) {
      if(data[i][0] && String(data[i][0]).toUpperCase() === cod) return data[i][1];
    }
    return "";
  } catch(e) { return ""; }
}

function obtenerListas() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let unidadesSheet = ss.getSheetByName(HOJA_UNIDADES);
    let gruposSheet = ss.getSheetByName(HOJA_GRUPOS);
    if (!unidadesSheet || !gruposSheet) return { unidades: [], grupos: [] };
    const unidadesData = unidadesSheet.getDataRange().getValues();
    const gruposData = gruposSheet.getDataRange().getValues();
    const unidades = unidadesData.slice(1).map(r => r[0]).filter(u => u);
    const grupos = gruposData.slice(1).map(r => r[0]).filter(g => g);
    return { unidades: unidades.sort(), grupos: grupos.sort() };
  } catch (error) { return { unidades: [], grupos: [] }; }
}

function obtenerResumen() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const prodSheet = ss.getSheetByName(HOJA_PRODUCTOS);
    const movSheet = ss.getSheetByName(HOJA_MOVIMIENTOS);
    if (!prodSheet || !movSheet) return { totalProductos: 0, totalMovimientos: 0, sinStock: 0, stockBajo: 0 };
    const productos = prodSheet.getDataRange().getValues();
    const movimientos = movSheet.getDataRange().getValues();
    const totalProductos = Math.max(0, productos.length - 1);
    const totalMovimientos = Math.max(0, movimientos.length - 1);
    let sinStock = 0, stockBajo = 0;
    for (let i = 1; i < productos.length; i++) {
      if (!productos[i][0]) continue;
      const codigo = productos[i][0];
      const stockMin = Math.max(0, parseInt(productos[i][4]) || 0);
      const stock = calcularStock(codigo);
      if (stock <= 0) sinStock++;
      else if (stock <= stockMin && stockMin > 0) stockBajo++;
    }
    return { totalProductos, totalMovimientos, sinStock, stockBajo };
  } catch (e) { return { totalProductos: 0, totalMovimientos: 0, sinStock: 0, stockBajo: 0 }; }
}

function obtenerMapaClientesPorClave() {
  const mapa = {};
  obtenerClientes().forEach(c => {
    const k = String(c.clave || '').toUpperCase();
    if (k) mapa[k] = c.nombre || '';
  });
  return mapa;
}

function obtenerMapaProductosPorCodigo() {
  const mapa = {};
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_PRODUCTOS);
    if (!sheet || sheet.getLastRow() < 2) return mapa;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) mapa[String(data[i][0]).toUpperCase()] = data[i][1] || '';
    }
    return mapa;
  } catch (e) { return mapa; }
}

function formatearFecha(fecha) {
  try {
    const f = new Date(fecha);
    return Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch (e) { return ""; }
}

/* ------------------- VALIDACIÓN E INICIALIZACIÓN ------------------- */
function validarIntegridad() {
  const errores = [];
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hojasRequeridas = [HOJA_PRODUCTOS, HOJA_MOVIMIENTOS, HOJA_UNIDADES, HOJA_GRUPOS, HOJA_CLIENTES, HOJA_KITS, HOJA_KIT_COMPONENTES, HOJA_CAJA, HOJA_PARAMETROS];
    hojasRequeridas.forEach(h => { if (!ss.getSheetByName(h)) errores.push(`Falta hoja: ${h}`); });
    return { errores };
  } catch (e) { return { errores: [e.message] }; }
}

function inicializarHojas() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const headersCaja19 = [[
      "IDREGISTRO","Fecha Captura","Cliente","Agente","Obra","Remisión","Tipo",
      "Producto","Cantidad","Fecha Renta Inicio","Fecha Renta Fin","Usuario","Timestamp","KitCodigo",
      "PrecioUnit","DescPct","DescMonto","SubTotalLinea","ImporteLinea"
    ]];

    const mapaHojas = {
      [HOJA_PRODUCTOS]: [["Código", "Nombre", "Unidad", "Grupo", "Stock Mínimo", "Fecha Creación"]],
      [HOJA_MOVIMIENTOS]: [["Código", "Fecha", "Tipo", "Cantidad", "Usuario", "Timestamp", "Observaciones", "Stock Resultante"]],
      [HOJA_CLIENTES]: [["Clave","Nombre","Teléfono","Email","Dirección","Estado","Fecha Creación"]],
      [HOJA_KITS]: [["CodigoKit","Nombre","Descripcion","Estado","Fecha Creacion"]],
      [HOJA_KIT_COMPONENTES]: [["CodigoKit","CodigoProducto","Cantidad"]],
      [HOJA_CAJA]: headersCaja19,
      [HOJA_UNIDADES]: [["Unidad"]],
      [HOJA_GRUPOS]: [["Grupo"]],
      [HOJA_PARAMETROS]: [[
        "Nombre", "RFC", "Dirección", "Web", "Colonia", "Ciudad", "Estado", "Pais", "CP", "Regimen", "Email", "Teléfono",
        "LogoFileId", "LogoUrl"
      ]]
    };

    let msgs = [];
    for (const [nombre, headers] of Object.entries(mapaHojas)) {
      let sheet = ss.getSheetByName(nombre);
      if (!sheet) {
        sheet = ss.insertSheet(nombre);
        sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
        msgs.push(`Creada hoja ${nombre}`);
      } else {
        if (nombre === HOJA_PARAMETROS && sheet.getLastColumn() < 14) {
          const lastCol = sheet.getLastColumn();
          sheet.insertColumnsAfter(lastCol, 14 - lastCol);
          sheet.getRange(1, 1, 1, 14).setValues(headers);
        }
        if (nombre === HOJA_CAJA) {
          const needCols = 19;
          const lastCol = sheet.getLastColumn();
          if (lastCol < needCols) sheet.insertColumnsAfter(lastCol, needCols - lastCol);
          sheet.getRange(1, 1, 1, needCols).setValues(headersCaja19);
        }
      }
    }

    return msgs.length ? msgs.join('\n') : "Todas las hojas ya existen.";
  } catch (e) {
    return "Error al inicializar: " + e.message;
  }
}