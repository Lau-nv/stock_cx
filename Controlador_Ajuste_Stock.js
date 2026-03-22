//esto es una prueba
function mostrarFormularioAjusteStock() {
  const html = HtmlService
    .createHtmlOutputFromFile('Formulario_Ajuste_Stock') // <- nombre del HTML SIN .html
    .setTitle('Ajuste de Stock')
    .setWidth(360); // sidebar
  SpreadsheetApp.getUi().showSidebar(html);
}

/************  LIBERACIONES  ************/

// Devuelve liberaciones pendientes (checkbox "Ingresado al Stock?" = FALSE)
// Formato: [{rowIndex, codigo, lote, cantidad}]
function obtenerLiberacionesPendientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Liberaciones');
  if (!hoja) throw new Error("No existe la hoja 'Liberaciones'.");

  const lastRow = hoja.getLastRow();
  const lastCol = hoja.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const headers = hoja.getRange(1, 1, 1, lastCol).getValues()[0].map(h => normaliza_(h));
  const idx = (needleArr) => {
    for (const needle of needleArr) {
      const pos = headers.findIndex(h => h.includes(needle));
      if (pos !== -1) return pos; // 0-based
    }
    return -1;
  };

  const COL_CODIGO = 8, COL_LOTE = 9, COL_CANT = 10, COL_CHECK = 13;
  // Buscamos columnas por nombre
  const cCodigo = idx(['codigo de producto','codigo producto','codigo']);
  const cLote   = COL_LOTE;
  const cCant   = idx(['cantidad por producto','cantidad']);
  const cCheck  = idx(['ingresado al stock','ingresado al stock?','ingresado']); // admite signos

  if (cCodigo === -1 || cLote === -1 || cCant === -1 || cCheck === -1) {
    throw new Error("Encabezados de 'Liberaciones' no encontrados (Código/Lote/Cantidad/Ingresado...). Revisá los nombres.");
  }

  const data = hoja.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const res = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const ingresado = !!row[cCheck];
    if (!ingresado) {
      const codigo = (row[cCodigo] || '').toString().trim().toUpperCase();
      const lote   = (row[cLote]   || '').toString().trim().toUpperCase();
      const cant   = Number(row[cCant] || 0);
      if (codigo && lote && cant > 0) {
        res.push({ rowIndex: i + 2, codigo, lote, cantidad: cant });
      }
    }
  }
  return res;
}

function normaliza_(s) {
  return (s || '').toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().trim();
}


// items: [{rowIndex, codigo, lote, cantidad}], observaciones: string (informativo)
function registrarIngresoDesdeLiberaciones(items, observaciones) {
  logInfo('Iniciando registrarIngresoDesdeLiberaciones', { itemsCount: items?.length || 0 });
  if (!items || !items.length) return "❌ No hay ítems para ingresar.";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName('Inventario');
  const liberSheet = ss.getSheetByName('Liberaciones');
  const movimientosSheet = ss.getSheetByName('Movimientos');
  if (!invSheet) throw new Error("No existe la hoja 'Inventario'.");
  if (!liberSheet) throw new Error("No existe la hoja 'Liberaciones'.");
  if (!movimientosSheet) throw new Error("No existe la hoja 'Movimientos'.");

  // Preparar arrays (solo ítems válidos)
  const rowsToMark = [];
  const movimientosRows = [];
  let totalIngresado = 0;

  const obsText = (observaciones || "LIBERACIÓN").trim() || "LIBERACIÓN";
  const ahora = new Date();
  const tipoMovimiento = "Ingreso desde Liberaciones";

  for (const it of items) {
    const COD = (it.codigo || '').toString().trim().toUpperCase();
    const LOT = (it.lote || '').toString().trim().toUpperCase();
    const QTY = Number(it.cantidad || 0);
    if (!COD || !LOT || !(QTY > 0)) continue;

    // Usar sumarEnInventario_ para manejar ID y logs correctamente
    // NOTA: No incluimos filaMovimiento porque los movimientos se registran en batch DESPUÉS
    // de todas las operaciones. Podría refactorizarse para registrar uno a uno.
    const datosLog = { tipoMovimiento, observaciones: obsText };
    const resultado = sumarEnInventario_(invSheet, COD, LOT, 'Depo', QTY, datosLog);
    
    if (resultado !== true) {
      logWarn('Error al sumar liberación en Inventario', { codigo: COD, lote: LOT, cantidad: QTY, error: resultado });
      continue; // Saltar este ítem y continuar con los demás
    }
    
    totalIngresado += QTY;
    rowsToMark.push(it.rowIndex);

    // Fila para Movimientos
    movimientosRows.push([
      ahora,                 // Fecha y Hora
      tipoMovimiento,        // Tipo de Movimiento
      COD,                   // Código Producto
      LOT,                   // Lote
      QTY,                   // Cantidad
      "N/A",                 // Caja Origen
      "N/A",                 // Caja Destino
      "N/A",                 // Paciente
      "N/A",                 // Cliente
      obsText,               // Observaciones
      0                      // EsConsumo
    ]);
  }

  const k = movimientosRows.length;
  if (k === 0) return "❌ No hay ítems válidos para ingresar.";

  // Agregar movimientos en bloque
  const movStart = movimientosSheet.getLastRow() + 1;
  movimientosSheet.getRange(movStart, 1, movimientosRows.length, 11).setValues(movimientosRows);

  // Marcar "Ingresado al Stock?" en TRUE para las filas de Liberaciones
  const COL_CHECK = 13;
  rowsToMark.forEach(r => liberSheet.getRange(r, COL_CHECK).setValue(true));

  SpreadsheetApp.flush();

  logInfo('Liberaciones ingresadas exitosamente', { itemsCount: k, totalUnidades: totalIngresado });
  return `✅ Ingresadas ${k} liberaciones (total ${totalIngresado} unidades) a Inventario y registradas en Movimientos.`;
}



/************  CAJAS DISPONIBLES (para UI)  ************/
function obtenerCajasDisponibles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName('Inventario');
  if (!invSheet) return [];
  
  const COL_UBIC = getColumnIndexInventario('ubicacion');
  const lastRow = invSheet.getLastRow();
  if (lastRow < 2) return [];
  
  const ubicaciones = invSheet.getRange(2, COL_UBIC, lastRow - 1, 1).getValues()
    .map(r => (r[0] || '').toString().trim().toUpperCase())
    .filter(v => v !== '' && v !== 'DEPO');
  
  return Array.from(new Set(ubicaciones)).sort((a,b)=>a.localeCompare(b,'es',{numeric:true}));
}

/************  AJUSTES: Ingreso / Egreso  ************/
function registrarStockRepo(datos) {
  logInfo('Iniciando registrarStockRepo', { datos });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName('Inventario');
  const movimientosSheet = ss.getSheetByName('Movimientos');

  if (!invSheet) return "❌ Falta la hoja 'Inventario'.";
  if (!movimientosSheet) return "❌ Falta la hoja 'Movimientos'.";

  // 🔧 Normalizar tipo de ajuste
  const tipoRaw = (datos.ajuste || "").toString();
  const tipo = tipoRaw
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().trim();

  // Campos comunes
  const codigo = (datos.codigo || "").toString().trim().toUpperCase();
  const lote   = (datos.lote   || "").toString().trim().toUpperCase();
  const observaciones = (datos.observaciones || "N/A").toString().trim() || "N/A";

  // Ajustes SIEMPRE piden cantidad
  let cantidad = parseInt(datos.cantidad, 10);
  if (isNaN(cantidad) || cantidad <= 0) {
    return "❌ Debes ingresar una cantidad válida.";
  }

  // Caja Origen/Destino a registrar en Movimientos (para trazabilidad)
  let cajaOrigen = "N/A";
  let cajaDestino = "N/A";

  // Validación básica
  if (!codigo || !lote) {
    return "❌ Debes completar Código y Lote.";
  }

  // ✅ Validar lote
  const validacion = validarLoteManual(codigo, lote);
  if (validacion !== true) {
    return validacion;
  }

  // Ejecutar según tipo
  let resultado = false;
  
  // 📝 Capturar número de fila ANTES de la operación para referencia en logs
  const filaMovimiento = movimientosSheet.getLastRow() + 1;

  if (tipo === "ingreso") {
    // Ingreso manual a Depo (Inventario con ubicación "Depo")
    const datosLog = { tipoMovimiento: 'Ingreso Manual', observaciones, filaMovimiento };
    resultado = (sumarEnInventario_(invSheet, codigo, lote, 'Depo', cantidad, datosLog) === true);

  } else if (tipo === "egreso") {
    // Egreso ⇒ desde Depo o desde Caja
    const origenTipo = (datos.origenTipo || "").toString();
    const origenCaja = (datos.origenCaja || "").toString().trim().toUpperCase();

    if (!origenTipo) return "❌ Indicá si el egreso es de Depo o de una Caja.";

    if (origenTipo === "Depo") {
      const datosLog = { tipoMovimiento: 'Egreso Manual', observaciones, filaMovimiento };
      const r = restarEnInventario_(invSheet, codigo, lote, 'Depo', cantidad, datosLog);
      if (r !== true) return r;
      resultado = true;
      cajaOrigen = "DEPO";
    } else if (origenTipo === "Caja") {
      if (!origenCaja) return "❌ Debés seleccionar la Caja Origen.";
      const datosLog = { tipoMovimiento: 'Egreso Manual', ubicacionOrigen: origenCaja, observaciones, filaMovimiento };
      const r = restarEnInventario_(invSheet, codigo, lote, origenCaja, cantidad, datosLog);
      if (r !== true) return r;
      resultado = true;
      cajaOrigen = origenCaja;
    } else {
      return "❌ Origen inválido. Elegí 'Depo' o 'Caja'.";
    }

  } else if (tipo === "ingreso desde liberaciones") {
    // Este flujo se confirma desde el modal (batch)
    return "ℹ️ Para 'Ingreso desde Liberaciones' usá el botón Confirmar del modal.";
  
  } else {
    return `❌ Tipo de ajuste no reconocido: "${tipoRaw}"`;
  }

  if (typeof resultado === "string") return resultado;

  if (resultado) {
    // 📝 Registrar movimiento en Movimientos (filaMovimiento ya se usó en los logs)
    movimientosSheet.appendRow([
      new Date(), tipoRaw, codigo, lote, cantidad,
      cajaOrigen, cajaDestino, "N/A", "N/A", observaciones, 0
    ]);

    // Limpieza & totales
    // limpiarCeros(['Inventario']); // ✅ Deshabilitado: ahora mantenemos lotes en 0 para trazabilidad
    if (typeof actualizarStockTotal === "function") actualizarStockTotal();

    logInfo('Ajuste registrado correctamente', { tipo: tipoRaw, codigo, lote, cantidad, filaMovimiento });
    return "✅ Movimiento registrado correctamente.";
  } else {
    logWarn('Ajuste no realizado', { tipo: tipoRaw, codigo, lote, resultado });
    return "❌ Movimiento no realizado. Verificá el stock y los datos ingresados.";
  }
}


/************  HELPERS (centralizados) ************/
// Las funciones concretas de manipulación de hojas (sumarEnRepo_, restarEnRepo_, sumarEnCajas_, restarEnCajas_, limpiarCeros, actualizarStockTotal, validarLoteManual, normalizarTipo_)
// se encuentran en Helpers.gs para evitar duplicados.



