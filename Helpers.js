/**
 * Helpers centralizados para gestión de stock
 * Contiene implementaciones canónicas para operaciones sobre Stock_Repo y Stock_Cajas,
 * validación de lotes, y utilidades compartidas.
 */

/*** SISTEMA DE LOGGING ***/
const LOG_LEVELS = { ERROR: 0, WARN: 1, INFO: 2, DEBUG: 3 };
const CURRENT_LOG_LEVEL = LOG_LEVELS.INFO; // Cambiar según necesidad

function logError(message, context = {}) {
  if (CURRENT_LOG_LEVEL >= LOG_LEVELS.ERROR) {
    Logger.log(`[ERROR] ${message} | Context: ${JSON.stringify(context)}`);
    console.error(message, context);
  }
}

function logWarn(message, context = {}) {
  if (CURRENT_LOG_LEVEL >= LOG_LEVELS.WARN) {
    Logger.log(`[WARN] ${message} | Context: ${JSON.stringify(context)}`);
  }
}

function logInfo(message, context = {}) {
  if (CURRENT_LOG_LEVEL >= LOG_LEVELS.INFO) {
    Logger.log(`[INFO] ${message} | Context: ${JSON.stringify(context)}`);
  }
}

function logDebug(message, context = {}) {
  if (CURRENT_LOG_LEVEL >= LOG_LEVELS.DEBUG) {
    Logger.log(`[DEBUG] ${message} | Context: ${JSON.stringify(context)}`);
  }
}

/*** SISTEMA DE AUDITORÍA ***/
function registrarAuditoria_(accion, datos, resultado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let auditSheet = ss.getSheetByName('Auditoria');
    
    // Crear hoja de auditoría si no existe
    if (!auditSheet) {
      auditSheet = ss.insertSheet('Auditoria');
      auditSheet.getRange(1, 1, 1, 6).setValues([
        ['Fecha y Hora', 'Usuario', 'Acción', 'Datos', 'Resultado', 'Detalles']
      ]);
      auditSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      auditSheet.hideSheet(); // Ocultar de usuarios normales
    }
    
    const usuario = Session.getActiveUser().getEmail() || 'usuario@desconocido.com';
    const datosStr = typeof datos === 'object' ? JSON.stringify(datos) : datos.toString();
    const resultadoStr = typeof resultado === 'object' ? JSON.stringify(resultado) : resultado.toString();
    
    auditSheet.appendRow([
      new Date(),
      usuario,
      accion,
      datosStr.substring(0, 500), // Limitar tamaño
      resultadoStr.substring(0, 200),
      resultado === true ? 'ÉXITO' : (typeof resultado === 'string' && resultado.includes('❌') ? 'ERROR' : 'ÉXITO')
    ]);
    
    logDebug('Auditoría registrada', { accion, usuario });
  } catch (e) {
    // No fallar si la auditoría falla
    logWarn('Error al registrar auditoría', { error: e.message });
  }
}

/*** SISTEMA DE VALIDACIÓN DE INTEGRIDAD ***/
function validarIntegridadInventario_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName('Inventario');
  
  if (!invSheet) {
    return { valido: true, errores: [], advertencias: [] };
  }
  
  const lastRow = invSheet.getLastRow();
  if (lastRow < 2) {
    return { valido: true, errores: [], advertencias: [] };
  }
  
  const COL_ID = getColumnIndexInventario('id');
  const COL_COD = getColumnIndexInventario('codigo');
  const COL_LOTE = getColumnIndexInventario('lote');
  const COL_CANT = getColumnIndexInventario('cantidad');
  const COL_UBIC = getColumnIndexInventario('ubicacion');
  
  const lastCol = invSheet.getLastColumn();
  const data = invSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  const errores = [];
  const advertencias = [];
  const idsVistos = new Map(); // Para detectar IDs duplicados
  
  for (let i = 0; i < data.length; i++) {
    const fila = i + 2;
    const id = (data[i][COL_ID - 1] || '').toString().trim();
    const codigo = (data[i][COL_COD - 1] || '').toString().trim();
    const lote = (data[i][COL_LOTE - 1] || '').toString().trim();
    const cantidad = Number(data[i][COL_CANT - 1] || 0);
    const ubicacion = (data[i][COL_UBIC - 1] || '').toString().trim();
    
    // Validación 1: Cantidad negativa (ERROR CRÍTICO)
    if (cantidad < 0) {
      errores.push(`Fila ${fila}: Cantidad negativa (${cantidad}) - Código: ${codigo}, Lote: ${lote}, Ubicación: ${ubicacion}`);
    }
    
    // Validación 2: Datos vacíos en fila con cantidad > 0
    if (cantidad > 0 && (!codigo || !lote || !ubicacion)) {
      errores.push(`Fila ${fila}: Datos incompletos - Código: "${codigo}", Lote: "${lote}", Ubicación: "${ubicacion}"`);
    }
    
    // Validación 3: Cantidad = 0 es VÁLIDA (se mantiene para trazabilidad con Logs)
    // Ya no se considera advertencia - las filas en 0 tienen valor histórico
    
    // Validación 4: ID faltante en filas con datos
    if (cantidad > 0 && !id) {
      advertencias.push(`Fila ${fila}: Falta ID único - se generará automáticamente en próximo movimiento`);
    }
    
    // Validación 5: ID duplicado (ahora el ID ya incluye ubicación)
    if (id) {
      if (idsVistos.has(id)) {
        errores.push(`Fila ${fila}: ID duplicado "${id}" (ya existe en fila ${idsVistos.get(id)})`);
      } else {
        idsVistos.set(id, fila);
      }
    }
    
    // Validación 6: ID no coincide con código+lote+ubicación
    if (id && codigo && lote && ubicacion) {
      const idEsperado = generarIDInventario_(codigo, lote, ubicacion);
      if (id !== idEsperado) {
        advertencias.push(`Fila ${fila}: ID "${id}" no coincide con código+lote+ubicación esperado "${idEsperado}"`);
      }
    }
  }
  
  const resultado = {
    valido: errores.length === 0,
    errores,
    advertencias,
    totalFilas: data.length,
    filasConErrores: errores.length,
    filasConAdvertencias: advertencias.length
  };
  
  if (!resultado.valido) {
    logError('Integridad de Inventario comprometida', resultado);
  } else if (advertencias.length > 0) {
    logWarn('Advertencias de integridad en Inventario', resultado);
  }
  
  return resultado;
}

/*** SISTEMA DE LÍMITES DE OPERACIÓN (CIRCUIT BREAKER) ***/
const LIMITES_OPERACION = {
  CANTIDAD_MAXIMA: 1000,           // Cantidad máxima por operación
  OPERACIONES_POR_MINUTO: 100,     // Rate limiting
  CANTIDAD_ALERTA: 500             // Advertencia si supera este umbral
};

function validarLimitesOperacion_(cantidad, contexto = {}) {
  const errores = [];
  const advertencias = [];
  
  // Validación 1: Cantidad máxima
  if (cantidad > LIMITES_OPERACION.CANTIDAD_MAXIMA) {
    errores.push(`❌ Cantidad ${cantidad} excede el límite máximo de ${LIMITES_OPERACION.CANTIDAD_MAXIMA} unidades por operación.`);
  }
  
  // Validación 2: Umbral de alerta
  if (cantidad > LIMITES_OPERACION.CANTIDAD_ALERTA && cantidad <= LIMITES_OPERACION.CANTIDAD_MAXIMA) {
    advertencias.push(`⚠️ ALERTA: Operación con cantidad elevada (${cantidad} unidades). Verificar que sea correcto.`);
  }
  
  // Validación 3: Rate limiting (operaciones por minuto)
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'ops_count_' + Session.getActiveUser().getEmail().replace(/[^a-zA-Z0-9]/g, '_');
    const contador = parseInt(cache.get(cacheKey) || '0');
    
    if (contador >= LIMITES_OPERACION.OPERACIONES_POR_MINUTO) {
      errores.push(`⏱️ Demasiadas operaciones en un minuto (${contador}). Esperá unos segundos antes de continuar.`);
    } else {
      cache.put(cacheKey, (contador + 1).toString(), 60); // Expira en 60 segundos
    }
  } catch (e) {
    // Si falla el cache, continuar (no bloquear la operación)
    logWarn('Error en rate limiting', { error: e.message });
  }
  
  const resultado = {
    valido: errores.length === 0,
    errores,
    advertencias,
    cantidad,
    contexto
  };
  
  if (!resultado.valido) {
    logError('Límites de operación excedidos', resultado);
    return resultado.errores[0]; // Retornar primer error
  }
  
  if (advertencias.length > 0) {
    logWarn('Advertencia en límites de operación', resultado);
  }
  
  return true;
}

/*** SISTEMA DE BLOQUEO (LOCKS) ***/
function executeWithLock_(operation, lockKey, timeoutMs = 30000) {
  const lock = LockService.getScriptLock();
  const startTime = new Date().getTime();
  
  try {
    // Intentar adquirir el lock con timeout
    const acquired = lock.tryLock(timeoutMs);
    if (!acquired) {
      const waitTime = new Date().getTime() - startTime;
      logError('No se pudo adquirir el lock', { lockKey, waitTime, timeoutMs });
      throw new Error(`⏱️ Operación bloqueada. Otro proceso está modificando el stock. Intentá nuevamente en unos segundos.`);
    }
    
    logDebug('Lock adquirido', { lockKey });
    const result = operation();
    logDebug('Operación completada con lock', { lockKey });
    return result;
    
  } catch (e) {
    logError('Error durante operación con lock', { lockKey, error: e.message, stack: e.stack });
    throw e;
  } finally {
    try {
      lock.releaseLock();
      logDebug('Lock liberado', { lockKey });
    } catch (e) {
      logWarn('Error al liberar lock', { lockKey, error: e.message });
    }
  }
}

/*** SISTEMA DE ROLLBACK TRANSACCIONAL ***/
function createSnapshot_(sheet, startRow, startCol, numRows, numCols) {
  try {
    if (!sheet || numRows <= 0 || numCols <= 0) return null;
    const data = sheet.getRange(startRow, startCol, numRows, numCols).getValues();
    return { sheet, startRow, startCol, data };
  } catch (e) {
    logWarn('No se pudo crear snapshot', { sheetName: sheet?.getName(), error: e.message });
    return null;
  }
}

function restoreSnapshot_(snapshot) {
  if (!snapshot || !snapshot.data) return false;
  try {
    snapshot.sheet.getRange(snapshot.startRow, snapshot.startCol, snapshot.data.length, snapshot.data[0].length)
      .setValues(snapshot.data);
    logInfo('Snapshot restaurado exitosamente', { sheetName: snapshot.sheet.getName() });
    return true;
  } catch (e) {
    logError('Error al restaurar snapshot', { sheetName: snapshot.sheet?.getName(), error: e.message });
    return false;
  }
}

// Normaliza el tipo (sin tildes, minúsculas, trim)
function normalizarTipo_(t) {
  return (t || '')
    .toString()
    .normalize('NFD').replace(/[\u0000-\u001F\u007F-\u007F\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

/*** SISTEMA DE VALIDACIONES CENTRALIZADAS ***/

// Valida que el código de producto exista en la hoja Productos
function validarCodigoProducto(codigo) {
  if (!codigo || codigo.trim() === '') {
    return '❌ El código de producto es obligatorio.';
  }
  
  const cod = codigo.toString().trim().toUpperCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productosSheet = ss.getSheetByName('Productos');
  
  if (!productosSheet) {
    logWarn('Hoja Productos no encontrada, saltando validación de código');
    return true; // Si no existe la hoja, no podemos validar pero permitimos continuar
  }
  
  const lastRow = productosSheet.getLastRow();
  if (lastRow < 2) {
    logWarn('Hoja Productos vacía, saltando validación');
    return true;
  }
  
  const codigos = productosSheet.getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(c => c.toString().trim().toUpperCase())
    .filter(c => c !== '');
  
  // Buscar coincidencia exacta o con wildcard
  const existe = codigos.some(codigoTabla => {
    if (codigoTabla.endsWith('*')) {
      const base = codigoTabla.slice(0, -1);
      return cod.startsWith(base);
    }
    return cod === codigoTabla;
  });
  
  if (!existe) {
    logWarn('Código no encontrado en Productos', { codigo: cod });
    return `❌ El código "${cod}" no existe en la hoja Productos.`;
  }
  
  logDebug('Código validado exitosamente', { codigo: cod });
  return true;
}

// Validación centralizada de datos de movimiento
function validarDatosMovimiento(datos) {
  const errores = [];
  
  // Validar código
  if (!datos.codigo || datos.codigo.trim() === '') {
    errores.push('Código de producto');
  }
  
  // Validar lote
  if (!datos.lote || datos.lote.trim() === '') {
    errores.push('Lote');
  }
  
  // Validar cantidad (si se requiere)
  if (datos.requiereCantidad && (!datos.cantidad || isNaN(parseInt(datos.cantidad)) || parseInt(datos.cantidad) <= 0)) {
    errores.push('Cantidad válida (debe ser un número mayor a 0)');
  }
  
  // Validar caja destino (si se requiere)
  if (datos.requiereCajaDestino && (!datos.cajaDestino || datos.cajaDestino.trim() === '' || datos.cajaDestino.trim().toUpperCase() === 'N/A')) {
    errores.push('Caja Destino');
  }
  
  // Validar caja origen (si se requiere)
  if (datos.requiereCajaOrigen && (!datos.cajaOrigen || datos.cajaOrigen.trim() === '' || datos.cajaOrigen.trim().toUpperCase() === 'N/A')) {
    errores.push('Caja Origen');
  }
  
  if (errores.length > 0) {
    return `❌ Campos faltantes o inválidos: ${errores.join(', ')}.`;
  }
  
  return true;
}

/*** SISTEMA DE ÍNDICES DINÁMICOS DE COLUMNAS ***/

// Obtiene índices de columnas por nombre (insensible a tildes y mayúsculas)
function obtenerIndicesColumnas(sheet, nombresColumnas) {
  if (!sheet) return null;
  
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return null;
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => normalizarTipo_(h));
  
  const indices = {};
  
  for (const [clave, variantes] of Object.entries(nombresColumnas)) {
    const variantesArray = Array.isArray(variantes) ? variantes : [variantes];
    const variantesNorm = variantesArray.map(v => normalizarTipo_(v));
    
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      // Usar coincidencia EXACTA primero, luego includes como fallback
      const matchExacto = variantesNorm.includes(header);
      const matchInclude = !matchExacto && variantesNorm.some(vn => 
        (vn.length > 2 && header.includes(vn)) || (header.length > 2 && vn.includes(header))
      );
      
      if (matchExacto || matchInclude) {
        indices[clave] = i + 1; // 1-based
        logDebug(`Columna '${clave}' encontrada en posición ${i + 1} (encabezado: "${sheet.getRange(1, i + 1).getValue()}")`);
        break;
      }
    }
  }
  
  return indices;
}

// Configuración de columnas esperadas para hoja Inventario unificada
const COLUMNAS_INVENTARIO = {
  id: ['id'],
  codigo: ['codigo', 'codigo de producto', 'codigo producto'],
  lote: ['lote', 'numero de lote'],
  cantidad: ['cantidad', 'cant', 'stock'],
  ubicacion: ['ubicacion', 'ubicación', 'location'],
  logs: ['logs', 'registro', 'historial']
};

// Cache de índices para evitar lecturas repetidas
let _cacheIndicesInventario = null;

function getColumnIndexInventario(columnKey) {
  if (!_cacheIndicesInventario) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Inventario');
    _cacheIndicesInventario = obtenerIndicesColumnas(sheet, COLUMNAS_INVENTARIO);
    
    // 🔍 Validación: Si no encontró TODAS las columnas, usar defaults
    const requiredKeys = ['id', 'codigo', 'lote', 'cantidad', 'ubicacion', 'logs'];
    const allFound = requiredKeys.every(key => _cacheIndicesInventario && _cacheIndicesInventario[key]);
    
    if (!allFound) {
      logWarn('No se encontraron todos los encabezados esperados en Inventario, usando índices por defecto', { 
        encontrados: _cacheIndicesInventario 
      });
      _cacheIndicesInventario = null; // Forzar uso de defaults
    }
  }
  return _cacheIndicesInventario?.[columnKey] || getDefaultColumnIndexInventario(columnKey);
}

// Fallback a índices hardcodeados si no se encuentran por nombre
function getDefaultColumnIndexInventario(columnKey) {
  const defaults = { id: 1, codigo: 2, lote: 3, cantidad: 4, ubicacion: 5, logs: 6 };
  return defaults[columnKey] || 1;
}

// Invalidar cache cuando sea necesario (ej: después de modificar estructura)
function invalidarCacheColumnas() {
  _cacheIndicesInventario = null;
  logDebug('Cache de índices de columnas invalidado');
}

// Generar ID único para inventario basado en código+lote+ubicación
function generarIDInventario_(codigo, lote, ubicacion) {
  const cod = (codigo || '').toString().trim().toUpperCase();
  const lot = (lote || '').toString().trim().toUpperCase();
  const ubi = (ubicacion || '').toString().trim().toUpperCase();
  return `${cod}_${lot}_${ubi}`;
}

// Generar entrada de log para movimientos
function generarLogMovimiento_(tipoMovimiento, cantidad, ubicacion, datos = {}) {
  const fecha = new Date();
  const tz = Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires';
  const fechaStr = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd HH:mm:ss');
  
  let detalles = '';
  if (datos.ubicacionOrigen && datos.ubicacionDestino) {
    detalles = ` (${datos.ubicacionOrigen}→${datos.ubicacionDestino})`;
  } else if (datos.ubicacionOrigen) {
    detalles = ` (desde ${datos.ubicacionOrigen})`;
  } else if (datos.ubicacionDestino) {
    detalles = ` (hacia ${datos.ubicacionDestino})`;
  }
  
  // Agregar referencia al movimiento si está disponible
  let refMovimiento = '';
  if (datos.filaMovimiento) {
    refMovimiento = ` [Mov. #${datos.filaMovimiento}]`;
  }
  
  return `${fechaStr} | ${tipoMovimiento}: ${cantidad > 0 ? '+' : ''}${cantidad} en ${ubicacion}${detalles}${refMovimiento}`;
}

// Agregar log a celda existente (separado por ";")
function agregarLogACelda_(logActual, nuevoLog) {
  if (!logActual || logActual.trim() === '') {
    return nuevoLog;
  }
  return `${logActual}; ${nuevoLog}`;
}

// VALIDACIÓN DE LOTES (versión canónica tomada y endurecida)
function validarLoteManual(codigo, lote) {
  const cod = (codigo || '').toString().trim().toUpperCase();
  const lot = (lote || '').toString().trim().toUpperCase();

  Logger.log('🔍 Código ingresado: ' + cod);
  Logger.log('🔍 Lote ingresado: ' + lot);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Prefijos');
  if (!hoja) return '❌ La hoja \'Prefijos\' no existe.';

  const data = hoja.getDataRange().getValues();
  let prefijoEsperado = null;

  for (let i = 1; i < data.length; i++) {
    const codigoTabla = (data[i][0] || '').toString().trim().toUpperCase();
    const prefijoTabla = (data[i][1] || '').toString().trim().toUpperCase();
    if (!codigoTabla || !prefijoTabla) continue;

    if (codigoTabla.endsWith('*')) {
      const base = codigoTabla.slice(0, -1);
      if (cod.startsWith(base)) {
        prefijoEsperado = prefijoTabla;
        Logger.log(`✅ Coincidencia parcial con "${codigoTabla}" → prefijo = "${prefijoEsperado}"`);
        break;
      }
    } else if (cod === codigoTabla) {
      prefijoEsperado = prefijoTabla;
      Logger.log(`✅ Coincidencia exacta con "${codigoTabla}" → prefijo = "${prefijoEsperado}"`);
      break;
    }
  }

  if (!prefijoEsperado) {
    Logger.log(`❌ No se encontró prefijo para código "${cod}"`);
    return `❌ No se encontró un prefijo válido para el código "${cod}" en la hoja Prefijos.`;
  }

  // Extraer posible parte de fecha
  const resto = lot.slice(prefijoEsperado.length);
  const fechaLote = resto.slice(0, 8);

  // Extraer año (defensivo)
  let anio = NaN;
  if (fechaLote.length >= 4) anio = parseInt(fechaLote.slice(0, 4));
  const esAnteriorA2023 = !isNaN(anio) && anio < 2023;

  if (!esAnteriorA2023 && !lot.startsWith(prefijoEsperado)) {
    Logger.log(`❌ El lote "${lot}" no comienza con el prefijo esperado "${prefijoEsperado}"`);
    return `❌ El lote debe comenzar con el prefijo "${prefijoEsperado}" según el código "${cod}".`;
  }

  const restoEvaluar = esAnteriorA2023 ? lot.replace(prefijoEsperado, '') : resto;
  if (!/^[0-9]{8}/.test(restoEvaluar)) {
    Logger.log('❌ El lote no tiene una fecha válida luego del prefijo');
    return '❌ Luego del prefijo, el lote debe comenzar con una fecha en formato AAAAMMDD.';
  }

  if (!/^[0-9]{15}$/.test(restoEvaluar)) {
    Logger.log(`❌ El lote luego del prefijo no tiene 15 dígitos válidos: "${restoEvaluar}"`);
    return '❌ Luego del prefijo, el lote debe tener exactamente 15 dígitos.';
  }

  Logger.log('✅ Lote válido con o sin prefijo según año, y con fecha y longitud correctos');
  return true;
}

/*** OPERACIONES EN INVENTARIO UNIFICADO ***/

/**
 * Suma cantidad en Inventario para un código+lote+ubicación
 * @param {Sheet} invSheet - Hoja Inventario
 * @param {string} codigo - Código del producto
 * @param {string} lote - Lote del producto
 * @param {string} ubicacion - Ubicación ("Depo" o nombre de caja)
 * @param {number} cantidad - Cantidad a sumar
 * @param {object} datosLog - Datos adicionales para el log (ubicacionOrigen, ubicacionDestino, etc.)
 * @returns {boolean|string} true si OK, string de error si falla
 */
function sumarEnInventario_(invSheet, codigo, lote, ubicacion, cantidad, datosLog = {}) {
  logInfo('Sumando en Inventario', { codigo, lote, ubicacion, cantidad });
  
  // ✅ VALIDACIÓN 1: Límites de operación
  const validacionLimites = validarLimitesOperacion_(cantidad, { accion: 'sumar', codigo, lote, ubicacion });
  if (validacionLimites !== true) {
    registrarAuditoria_('SUMAR_INVENTARIO', { codigo, lote, ubicacion, cantidad }, validacionLimites);
    return validacionLimites;
  }
  
  const COL_ID = getColumnIndexInventario('id');
  const COL_COD = getColumnIndexInventario('codigo');
  const COL_LOTE = getColumnIndexInventario('lote');
  const COL_CANT = getColumnIndexInventario('cantidad');
  const COL_UBIC = getColumnIndexInventario('ubicacion');
  const COL_LOGS = getColumnIndexInventario('logs');
  
  // 🔍 Debug: Verificar índices de columnas
  logDebug('Índices de columnas en sumarEnInventario_', { 
    COL_ID, COL_COD, COL_LOTE, COL_CANT, COL_UBIC, COL_LOGS 
  });
  
  const idUnico = generarIDInventario_(codigo, lote, ubicacion);
  const last = invSheet.getLastRow();
  const lastCol = invSheet.getLastColumn();
  
  // Buscar si ya existe el registro (código+lote+ubicación)
  if (last > 1) {
    const vals = invSheet.getRange(2, 1, last - 1, lastCol).getValues();
    for (let i = 0; i < vals.length; i++) {
      const idExistente = (vals[i][COL_ID - 1] || '').toString().trim();
      const cod = (vals[i][COL_COD - 1] || '').toString().trim().toUpperCase();
      const lot = (vals[i][COL_LOTE - 1] || '').toString().trim().toUpperCase();
      const ubic = (vals[i][COL_UBIC - 1] || '').toString().trim().toUpperCase();
      
      // Verificar por ID (ahora incluye ubicación) o por código+lote+ubicación
      if (idExistente === idUnico || (cod === codigo && lot === lote && ubic === ubicacion.toUpperCase())) {
        const fila = i + 2;
        const antigua = Number(vals[i][COL_CANT - 1] || 0);
        const nueva = antigua + cantidad;
        const logsActuales = (vals[i][COL_LOGS - 1] || '').toString();
        
        // Generar log del movimiento - usar tipo específico o fallback
        const tipoMovimiento = datosLog.tipoMovimiento || (cantidad > 0 ? 'Suma' : 'Resta');
        const nuevoLog = generarLogMovimiento_(tipoMovimiento, cantidad, ubicacion, datosLog);
        const logsActualizados = agregarLogACelda_(logsActuales, nuevoLog);
        
        logDebug('Actualizando fila existente en Inventario', { fila, ubicacion, antigua, nueva });
        
        // Actualizar cantidad y logs
        invSheet.getRange(fila, COL_CANT).setValue(nueva);
        invSheet.getRange(fila, COL_LOGS).setValue(logsActualizados);
        
        // Asegurar que ID esté presente (por si es un registro antiguo)
        if (!idExistente || idExistente === '') {
          invSheet.getRange(fila, COL_ID).setValue(idUnico);
        }
        
        SpreadsheetApp.flush();
        
        // ✅ Auditoría exitosa
        registrarAuditoria_('SUMAR_INVENTARIO', { codigo, lote, ubicacion, cantidad, antigua, nueva, fila }, true);
        return true;
      }
    }
  }
  
  // No existe, crear nueva fila
  const fecha = new Date();
  const tz = Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires';
  const fechaStr = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd HH:mm:ss');
  const logInicial = `${fechaStr} | Cantidad inicial: ${cantidad} en ${ubicacion}`;
  
  // ⚠️ IMPORTANTE: Crear array con estructura fija [ID, Codigo, Lote, Cantidad, Ubicacion, Logs]
  // NO depender de COL_* que pueden estar incorrectos
  const nuevaFila = [
    idUnico,      // Columna A: ID
    codigo,       // Columna B: Codigo  
    lote,         // Columna C: Lote
    cantidad,     // Columna D: Cantidad
    ubicacion,    // Columna E: Ubicacion
    logInicial    // Columna F: Logs
  ];
  
  const filaNum = last + 1;
  logDebug('Agregando nueva fila en Inventario', { fila: filaNum, codigo, lote, ubicacion, cantidad, estructura: nuevaFila });
  invSheet.appendRow(nuevaFila);
  SpreadsheetApp.flush();
  
  // ✅ Auditoría exitosa
  registrarAuditoria_('SUMAR_INVENTARIO_NUEVO', { codigo, lote, ubicacion, cantidad, fila: filaNum }, true);
  return true;
}

/**
 * Resta cantidad en Inventario para un código+lote+ubicación
 * @param {Sheet} invSheet - Hoja Inventario
 * @param {string} codigo - Código del producto
 * @param {string} lote - Lote del producto
 * @param {string} ubicacion - Ubicación ("Depo" o nombre de caja)
 * @param {number} cantidad - Cantidad a restar
 * @param {object} datosLog - Datos adicionales para el log (ubicacionOrigen, ubicacionDestino, etc.)
 * @returns {boolean|string} true si OK, string de error si falla
 */
function restarEnInventario_(invSheet, codigo, lote, ubicacion, cantidad, datosLog = {}) {
  logInfo('Restando en Inventario', { codigo, lote, ubicacion, cantidad });
  
  // ✅ VALIDACIÓN 1: Límites de operación
  const validacionLimites = validarLimitesOperacion_(cantidad, { accion: 'restar', codigo, lote, ubicacion });
  if (validacionLimites !== true) {
    registrarAuditoria_('RESTAR_INVENTARIO', { codigo, lote, ubicacion, cantidad }, validacionLimites);
    return validacionLimites;
  }
  
  const COL_ID = getColumnIndexInventario('id');
  const COL_COD = getColumnIndexInventario('codigo');
  const COL_LOTE = getColumnIndexInventario('lote');
  const COL_CANT = getColumnIndexInventario('cantidad');
  const COL_UBIC = getColumnIndexInventario('ubicacion');
  const COL_LOGS = getColumnIndexInventario('logs');
  
  // 🔍 Debug: Verificar índices de columnas
  logDebug('Índices de columnas en restarEnInventario_', { 
    COL_ID, COL_COD, COL_LOTE, COL_CANT, COL_UBIC, COL_LOGS 
  });
  
  const last = invSheet.getLastRow();
  const lastCol = invSheet.getLastColumn();
  
  if (last < 2) {
    logWarn('No hay registros en Inventario');
    const error = '❌ No hay registros en Inventario.';
    registrarAuditoria_('RESTAR_INVENTARIO_VACIO', { codigo, lote, ubicacion, cantidad }, error);
    return error;
  }

  const dataRange = invSheet.getRange(2, 1, last - 1, lastCol);
  const data = dataRange.getValues();
  let restante = cantidad;

  // 1) Calcular disponibilidad total para este código+lote+ubicación
  let totalDisponible = 0;
  for (let i = 0; i < data.length; i++) {
    const ok = (data[i][COL_COD - 1] || '').toString().trim().toUpperCase() === codigo
            && (data[i][COL_LOTE - 1] || '').toString().trim().toUpperCase() === lote
            && (data[i][COL_UBIC - 1] || '').toString().trim().toUpperCase() === ubicacion.toUpperCase();
    if (ok) totalDisponible += Number(data[i][COL_CANT - 1] || 0);
  }
  
  // ✅ VALIDACIÓN 2: Stock suficiente
  if (totalDisponible < cantidad) {
    logWarn('Stock insuficiente en Inventario', { ubicacion, codigo, lote, disponible: totalDisponible, solicitado: cantidad });
    const error = `❌ Stock insuficiente en ${ubicacion}. Disponible: ${totalDisponible}, a restar: ${cantidad}.`;
    registrarAuditoria_('RESTAR_INVENTARIO_INSUFICIENTE', { codigo, lote, ubicacion, cantidad, totalDisponible }, error);
    return error;
  }

  // 2) Restar de las filas que matcheen (puede haber duplicados)
  const filasActualizadas = []; // Para rastrear filas que necesitan actualizar logs
  
  for (let i = 0; i < data.length && restante > 0; i++) {
    const ok = (data[i][COL_COD - 1] || '').toString().trim().toUpperCase() === codigo
            && (data[i][COL_LOTE - 1] || '').toString().trim().toUpperCase() === lote
            && (data[i][COL_UBIC - 1] || '').toString().trim().toUpperCase() === ubicacion.toUpperCase();
    if (!ok) continue;
    
    const cant = Number(data[i][COL_CANT - 1] || 0);
    if (cant <= 0) continue;

    let cantidadRestada = 0;
    if (cant <= restante) {
      cantidadRestada = cant;
      data[i][COL_CANT - 1] = 0;
      restante -= cant;
      logDebug('Fila en Inventario reducida a 0', { fila: i + 2, cantidad: cant });
    } else {
      cantidadRestada = restante;
      data[i][COL_CANT - 1] = cant - restante;
      logDebug('Fila en Inventario actualizada', { fila: i + 2, antigua: cant, nueva: data[i][COL_CANT - 1] });
      restante = 0;
    }
    
    // Generar log de la resta - usar tipo específico o fallback
    const tipoMovimiento = datosLog.tipoMovimiento || 'Resta';
    const nuevoLog = generarLogMovimiento_(tipoMovimiento, -cantidadRestada, ubicacion, datosLog);
    const logsActuales = (data[i][COL_LOGS - 1] || '').toString();
    data[i][COL_LOGS - 1] = agregarLogACelda_(logsActuales, nuevoLog);
    
    filasActualizadas.push({ fila: i + 2, logs: data[i][COL_LOGS - 1] });
  }

  // 3) Aplicar cambios: actualizar cantidad y logs
  for (let i = data.length - 1; i >= 0; i--) {
    const rowIdx = i + 2;
    const nuevaCant = Number(data[i][COL_CANT - 1] || 0);
    
    // Verificar si esta fila fue modificada
    const filaActualizada = filasActualizadas.find(f => f.fila === rowIdx);
    
    if (filaActualizada) {
      // Esta fila fue modificada - actualizar cantidad y logs
      invSheet.getRange(rowIdx, COL_CANT).setValue(nuevaCant);
      invSheet.getRange(rowIdx, COL_LOGS).setValue(filaActualizada.logs);
      
      if (nuevaCant === 0) {
        logInfo('Fila reducida a cantidad=0 en Inventario (logs preservados)', { fila: rowIdx });
      } else {
        logDebug('Fila actualizada en Inventario', { fila: rowIdx, nuevaCantidad: nuevaCant });
      }
    }
  }

  SpreadsheetApp.flush();
  logDebug('Operación restarEnInventario_ completada exitosamente');
  
  // ✅ Auditoría exitosa
  registrarAuditoria_('RESTAR_INVENTARIO', { codigo, lote, ubicacion, cantidad, totalDisponible }, true);
  return true;
}

// Funciones de compatibilidad (wrappers) para código existente
function sumarEnRepo_(repoSheet, codigo, lote, cantidad) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName('Inventario');
  return sumarEnInventario_(invSheet, codigo, lote, 'Depo', cantidad);
}

function restarEnRepo_(repoSheet, codigo, lote, cantidad) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName('Inventario');
  return restarEnInventario_(invSheet, codigo, lote, 'Depo', cantidad);
}

function sumarEnCajas_(cajasSheet, codigo, lote, caja, cantidad) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName('Inventario');
  return sumarEnInventario_(invSheet, codigo, lote, caja, cantidad);
}

function restarEnCajas_(cajasSheet, codigo, lote, caja, cantidad) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName('Inventario');
  return restarEnInventario_(invSheet, codigo, lote, caja, cantidad);
}

/*** UTILIDADES ***/
// Versión OPTIMIZADA con escritura batch única - ahora para hoja Inventario unificada
function limpiarCeros(hojaNombres = ['Inventario']) {
  logInfo('Iniciando limpiarCeros', { hojas: hojaNombres });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Procesar la hoja Inventario
  hojaNombres.forEach(nombre => {
    const sh = ss.getSheetByName(nombre);
    if (!sh) {
      logWarn('Hoja no encontrada para limpiarCeros', { hoja: nombre });
      return;
    }
    
    const COL_CANT = getColumnIndexInventario('cantidad');
    
    const lastCol = sh.getLastColumn();
    const last = sh.getLastRow();
    if (last < 2) return; // nada que hacer

    const dataRowsCount = last - 1;

    // 1) Leer todos los datos en memoria (UNA SOLA LECTURA)
    const allData = sh.getRange(2, 1, dataRowsCount, lastCol).getValues();

    // 2) Filtrar filas con cantidad > 0
    const remaining = allData.filter(row => Number(row[COL_CANT - 1] || 0) > 0);

    if (remaining.length === dataRowsCount) {
      // No hay cambios necesarios
      logDebug('No hay filas con cantidad=0 para limpiar', { hoja: nombre });
      return;
    }

    // 3) OPTIMIZACIÓN: Escribir toda la data filtrada de una sola vez
    if (remaining.length > 0) {
      // Escribir datos filtrados
      sh.getRange(2, 1, remaining.length, lastCol).setValues(remaining);
      
      // Limpiar filas sobrantes (si quedan)
      const filasALimpiar = dataRowsCount - remaining.length;
      if (filasALimpiar > 0) {
        sh.getRange(2 + remaining.length, 1, filasALimpiar, lastCol).clearContent();
      }
    } else {
      // No quedan filas válidas, limpiar todo
      sh.getRange(2, 1, dataRowsCount, lastCol).clearContent();
    }

    logInfo('LimpiarCeros completado', { hoja: nombre, filasOriginales: dataRowsCount, filasRestantes: remaining.length });
  });
  
  SpreadsheetApp.flush();
  logDebug('limpiarCeros finalizado para todas las hojas');
  
  // ✅ VALIDACIÓN 3: Verificar integridad post-limpieza
  try {
    const validacion = validarIntegridadInventario_();
    if (!validacion.valido) {
      logWarn('Advertencia de integridad post-limpieza', { errores: validacion.errores, advertencias: validacion.advertencias });
      registrarAuditoria_('LIMPIAR_CEROS_VALIDACION', { 
        errores: validacion.errores.length, 
        advertencias: validacion.advertencias.length 
      }, validacion);
    } else {
      logInfo('Validación de integridad post-limpieza exitosa', { totalFilas: validacion.totalFilas });
    }
  } catch (e) {
    logError('Error al validar integridad post-limpieza', { error: e.toString() });
  }
}

// Actualizar Stock_Total - ahora lee solo de Inventario y consolida por ubicación
function actualizarStockTotal() {
  logInfo('Iniciando actualizarStockTotal');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName('Inventario');
  const totalSheet = ss.getSheetByName('Stock_Total');
  
  if (!totalSheet) {
    logWarn('Hoja Stock_Total no encontrada');
    return;
  }

  // Limpiar contenido anterior manteniendo encabezado
  const maxRows = totalSheet.getMaxRows();
  if (maxRows > 1) {
    const colsToClear = Math.max(totalSheet.getLastColumn(), 5);
    totalSheet.getRange(2, 1, Math.max(1, maxRows - 1), colsToClear).clearContent();
  }
  
  // Asegurar que el encabezado esté presente
  totalSheet.getRange(1, 1, 1, 5).setValues([['Código', 'Lote', 'Cantidad Depo', 'Cantidad Cajas', 'Total Stock']]);

  if (!invSheet) {
    logWarn('Hoja Inventario no encontrada');
    return;
  }

  const invLastRow = invSheet.getLastRow();
  if (invLastRow < 2) {
    logDebug('Inventario vacío, nada que consolidar');
    return;
  }

  const COL_COD = getColumnIndexInventario('codigo');
  const COL_LOTE = getColumnIndexInventario('lote');
  const COL_CANT = getColumnIndexInventario('cantidad');
  const COL_UBIC = getColumnIndexInventario('ubicacion');

  const lastCol = invSheet.getLastColumn();
  const invData = invSheet.getRange(2, 1, invLastRow - 1, lastCol).getValues();

  // Consolidar por código+lote, separando Depo vs Cajas
  const stockMap = {};
  for (const row of invData) {
    const codigo = (row[COL_COD - 1] || '').toString().trim().toUpperCase();
    const lote = (row[COL_LOTE - 1] || '').toString().trim().toUpperCase();
    const cantidad = Number(row[COL_CANT - 1] || 0);
    const ubicacion = (row[COL_UBIC - 1] || '').toString().trim().toUpperCase();
    
    if (!codigo || !lote || cantidad <= 0) continue;
    
    const key = `${codigo}_${lote}`;
    if (!stockMap[key]) {
      stockMap[key] = { codigo, lote, depo: 0, cajas: 0 };
    }
    
    if (ubicacion === 'DEPO') {
      stockMap[key].depo += cantidad;
    } else {
      stockMap[key].cajas += cantidad;
    }
  }

  // Construir todas las filas en memoria y escribirlas en batch
  const rows = [];
  for (const item of Object.values(stockMap)) {
    const depoNum = Number(item.depo || 0);
    const cajasNum = Number(item.cajas || 0);
    const total = depoNum + cajasNum;
    if (total > 0) {
      rows.push([item.codigo, item.lote, depoNum, cajasNum, total]);
    }
  }

  if (rows.length > 0) {
    // Asegurar que la hoja tiene suficientes filas
    const requiredRows = rows.length + 1;
    if (maxRows < requiredRows) {
      totalSheet.insertRowsAfter(maxRows, requiredRows - maxRows);
    }

    // Escritura en una sola llamada
    totalSheet.getRange(2, 1, rows.length, 5).setValues(rows);
    logInfo('Stock_Total actualizado', { registros: rows.length });
  } else {
    logDebug('No hay datos para escribir en Stock_Total');
  }
}
