
function mostrarFormularioQR() {
  const html = HtmlService.createHtmlOutputFromFile('FormularioQR')
    .setWidth(400)
    .setTitle("Registro de Movimientos");
  SpreadsheetApp.getUi().showSidebar(html);
}

function obtenerCodigosProductos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productosSheet = ss.getSheetByName('Productos');
  const lastRow = productosSheet.getLastRow();
  const productosData = productosSheet.getRange("A2:A" + lastRow).getValues().flat().filter(c => c.trim() !== "");
  return [...new Set(productosData)];
}

function registrarMovimiento(datos) {
  logInfo('Iniciando registrarMovimiento', { datos });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName('Inventario');
  const movimientosSheet = ss.getSheetByName('Movimientos');

  // 🔧 Normalizar tipo
  const tipoRaw = (datos.tipoMovimiento || "").toString();
  const tipo = tipoRaw
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // quita tildes
    .toLowerCase()
    .trim();

  const codigo = datos.codigo.trim().toUpperCase();
  const paciente = datos.paciente?.trim() || "N/A";
  const cliente = datos.cliente?.trim() || "N/A";
  const lote = datos.lote.trim().toUpperCase();

  // ✅ Validar código de producto
  const validacionCodigo = validarCodigoProducto(codigo);
  if (validacionCodigo !== true) {
    return validacionCodigo;
  }

  // ✅ Parsear cantidad (por defecto 1 si no viene o es inválida)
  let cantidad = parseInt(datos.cantidad) || 1;
  if (cantidad <= 0) {
    cantidad = 1;
  }

  const cajaOrigen = (datos.caja_origen?.trim() || "N/A").toUpperCase();
  const cajaDestino = (datos.caja_destino?.trim() || "N/A").toUpperCase();
  const observaciones = datos.observaciones?.trim() || "N/A";

  // ✅ Validación centralizada de datos básicos
  const validacionDatos = validarDatosMovimiento({
    codigo,
    lote,
    cantidad,
    requiereCantidad: true, // Todos los movimientos requieren cantidad
    requiereCajaDestino: (tipo === "reposicion" || tipo === "reposicion caja completa"),
    cajaDestino,
    requiereCajaOrigen: (tipo === "entre cajas" || tipo === "consumo"),
    cajaOrigen
  });
  
  if (validacionDatos !== true) {
    return validacionDatos;
  }

  // ✅ Validar lote
  const validacionLote = validarLoteManual(codigo, lote);
  if (validacionLote !== true) {
    logWarn("Validación de lote fallida", { codigo, lote, resultado: validacionLote });
    return validacionLote;
  }

  // 📝 Obtener número de fila donde se registrará en Movimientos (para referencia en logs)
  const filaMovimiento = movimientosSheet.getLastRow() + 1;
  
  let stockActualizado = false;

  if (tipo === "reposicion") {
    stockActualizado = moverDeRepoACajas(invSheet, codigo, lote, cantidad, cajaDestino, filaMovimiento);
  } else if (tipo === "reposicion caja completa") {
    // Igual que Reposición, pero permite caja nueva
    stockActualizado = moverDeRepoACajas(invSheet, codigo, lote, cantidad, cajaDestino, filaMovimiento);
  } else if (tipo === "entre cajas") {
    stockActualizado = moverEntreCajas(invSheet, codigo, lote, cantidad, cajaOrigen, cajaDestino, filaMovimiento);
  } else if (tipo === "consumo") {
    stockActualizado = moverDeCajasAConsumo(invSheet, codigo, lote, cantidad, cajaOrigen, filaMovimiento);
  } else if (tipo === "distribucion") {
    // Distribución: el origen puede ser Depo o una Caja
    const origenTipo = (datos.origenTipo || "").toString(); // 'Depo' | 'Caja'
    const origenCaja = (datos.caja_origen || "").toString().trim().toUpperCase();
    if (!origenTipo) {
      return "❌ Indicá si la distribución sale de Depo o de una Caja.";
    }
    if (origenTipo === 'Depo') {
      stockActualizado = moverDeRepoADistribucion(invSheet, codigo, lote, cantidad, filaMovimiento);
    } else if (origenTipo === 'Caja') {
      if (!origenCaja) return "❌ Debés seleccionar la Caja Origen.";
      const datosLog = { tipoMovimiento: 'Distribución', filaMovimiento };
      const r = restarEnInventario_(invSheet, codigo, lote, origenCaja, cantidad, datosLog);
      if (r !== true) return r;
      stockActualizado = true;
    } else {
      return "❌ Origen inválido. Elegí 'Depo' o 'Caja'.";
    }
  } else {
    // 🔴 Diagnóstico claro si no matchea
    return `❌ Tipo de movimiento no reconocido: "${tipoRaw}"`;
  }

  if (typeof stockActualizado === "string") {
    logWarn('Movimiento no realizado, retornando mensaje de error', { mensaje: stockActualizado });
    return stockActualizado;
  }

  if (stockActualizado) {
    // Registrar en Movimientos DESPUÉS de actualización exitosa de Inventario
    const filaMovimiento = movimientosSheet.getLastRow() + 1;
    const idCx = (tipo === "consumo") ? (datos.idCx?.toString().trim() || "N/A") : "N/A";
    movimientosSheet.appendRow([new Date(), tipoRaw, codigo, lote, cantidad, cajaOrigen, cajaDestino, paciente, cliente, observaciones, idCx]);
    
    // limpiarCeros(['Inventario']); // ✅ Deshabilitado: ahora mantenemos lotes en 0 para trazabilidad
    actualizarStockTotal();
    if (tipo === "consumo") generarReporteCX(idCx);
    logInfo('Movimiento registrado correctamente', { tipo: tipoRaw, codigo, lote, cantidad, filaMovimiento });
    return "✅ Movimiento registrado correctamente.";
  } else {
    // Si alguna vez llega acá, ya sabés que entró a una rama pero devolvió falsy.
    // moverDeRepoACajas/moverDe... devuelven true o string; falsy sugiere otra cosa.
    logError('Movimiento no realizado (valor falsy)', { tipo: tipoRaw, codigo, lote, stockActualizado });
    return "❌ Movimiento no realizado. Stock insuficiente del lote seleccionado.";
  }
}

// 🔀 Mover implantes entre cajas
// Valida existencia de (Código+Lote) en caja origen y stock suficiente.
// Ahora trabaja con la hoja Inventario unificada
function obtenerPacientesAgenda() {
  try {
    const agendaSS = SpreadsheetApp.openById('1Kg3J6dTS2SaUvz5AhDB8i6hgfLrUO_E9j43ymNxrxoM');
    const hojas = agendaSS.getSheets().filter(h =>
      h.getName().toLowerCase().startsWith('agenda')
    );
    const nombres = new Set();
    hojas.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      // Columna D = Paciente (columna 4, índice 3)
      sheet.getRange(2, 4, lastRow - 1, 1).getValues().forEach(row => {
        const nombre = (row[0] || '').toString().trim();
        if (nombre) nombres.add(nombre);
      });
    });
    return [...nombres].sort((a, b) => a.localeCompare(b, 'es'));
  } catch (e) {
    logError('Error obteniendo pacientes de agenda', { error: e.message });
    return [];
  }
}

function buscarIdCxPorPaciente(nombre) {
  if (!nombre || nombre.trim() === "") return null;
  const normalizar = s => s.toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  const nombreNorm = normalizar(nombre);
  try {
    const agendaSS = SpreadsheetApp.openById('1Kg3J6dTS2SaUvz5AhDB8i6hgfLrUO_E9j43ymNxrxoM');
    const hojas = agendaSS.getSheets().filter(h =>
      h.getName().toLowerCase().startsWith('agenda')
    );
    const coincidencias = [];
    hojas.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      // Columna C = ID (índice 0 del rango), Columna D = Paciente (índice 1)
      const data = sheet.getRange(2, 3, lastRow - 1, 2).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        const id = (data[i][0] || '').toString().trim();
        const paciente = (data[i][1] || '').toString().trim();
        if (id && paciente && normalizar(paciente) === nombreNorm) {
          coincidencias.push(id);
        }
      }
    });
    return coincidencias.length > 0 ? coincidencias : null;
  } catch (e) {
    logError('Error buscando ID CX en agenda', { error: e.message });
    return null;
  }
}

function moverEntreCajas(invSheet, codigo, lote, cantidad, cajaOrigen, cajaDestino, filaMovimiento) {
  // Validaciones básicas
  if (!invSheet) throw new Error('Se requiere la hoja Inventario.');
  if (!codigo || !lote || !cajaOrigen || !cajaDestino || isNaN(cantidad) || cantidad <= 0) {
    return "❌ Datos inválidos. Revisá código, lote, cajas y cantidad.";
  }
  if (cajaOrigen === cajaDestino) {
    return "❌ La caja de origen y de destino no pueden ser la misma.";
  }

  // Ejecutar con lock para evitar condiciones de carrera
  return executeWithLock_(() => {
    logInfo('Iniciando moverEntreCajas', { codigo, lote, cantidad, cajaOrigen, cajaDestino });
    
    // Crear snapshot para rollback
    const invSnapshot = createSnapshot_(invSheet, 2, 1, invSheet.getLastRow() - 1, invSheet.getLastColumn());

    try {
      // 1) Restar de la caja origen
      const datosLogResta = { 
        tipoMovimiento: 'Movimiento',
        ubicacionOrigen: cajaOrigen, 
        ubicacionDestino: cajaDestino,
        filaMovimiento 
      };
      const resultadoOrigen = restarEnInventario_(invSheet, codigo, lote, cajaOrigen, cantidad, datosLogResta);
      if (resultadoOrigen !== true) {
        logWarn('Falló resta en caja origen', { resultado: resultadoOrigen });
        return resultadoOrigen;
      }

      // 2) Sumar en la caja destino
      const datosLogSuma = { 
        tipoMovimiento: 'Movimiento',
        ubicacionOrigen: cajaOrigen, 
        ubicacionDestino: cajaDestino,
        filaMovimiento 
      };
      const resultadoDestino = sumarEnInventario_(invSheet, codigo, lote, cajaDestino, cantidad, datosLogSuma);
      if (resultadoDestino !== true) {
        // Rollback: restaurar Inventario
        logError('Falló suma en caja destino, iniciando rollback', { resultado: resultadoDestino });
        if (invSnapshot) {
          restoreSnapshot_(invSnapshot);
          logInfo('Rollback de Inventario completado');
        } else {
          logError('No se pudo hacer rollback, snapshot de Inventario no disponible');
        }
        return resultadoDestino;
      }

      logInfo('moverEntreCajas completado exitosamente');
      return true;
      
    } catch (e) {
      // Rollback en caso de excepción
      logError('Excepción en moverEntreCajas, iniciando rollback', { error: e.message, stack: e.stack });
      if (invSnapshot) restoreSnapshot_(invSnapshot);
      return `❌ Error crítico: ${e.message}`;
    }
  }, 'moverEntreCajas');
}


function moverDeRepoACajas(invSheet, codigo, lote, cantidad, cajaDestino, filaMovimiento) {
  // Validaciones básicas
  if (!invSheet) throw new Error('Se requiere la hoja Inventario.');
  if (!codigo || !lote || !cajaDestino || isNaN(cantidad) || cantidad <= 0) {
    return "❌ Datos inválidos. Revisa código, lote, caja destino y cantidad.";
  }

  // Ejecutar con lock para evitar condiciones de carrera
  return executeWithLock_(() => {
    logInfo('Iniciando moverDeRepoACajas', { codigo, lote, cantidad, cajaDestino });
    
    // Crear snapshot para rollback
    const invSnapshot = createSnapshot_(invSheet, 2, 1, invSheet.getLastRow() - 1, invSheet.getLastColumn());

    try {
      // 1) Restar de Depo
      const datosLogResta = { 
        tipoMovimiento: 'Reponer',
        ubicacionOrigen: 'Depo', 
        ubicacionDestino: cajaDestino,
        filaMovimiento 
      };
      const resultadoDepo = restarEnInventario_(invSheet, codigo, lote, 'Depo', cantidad, datosLogResta);
      if (resultadoDepo !== true) {
        logWarn('Falló resta en Depo', { resultado: resultadoDepo });
        return resultadoDepo;
      }

      // 2) Sumar en Caja
      const datosLogSuma = { 
        tipoMovimiento: 'Reponer',
        ubicacionOrigen: 'Depo', 
        ubicacionDestino: cajaDestino,
        filaMovimiento 
      };
      const resultadoCaja = sumarEnInventario_(invSheet, codigo, lote, cajaDestino, cantidad, datosLogSuma);
      if (resultadoCaja !== true) {
        // Rollback: restaurar Inventario
        logError('Falló suma en caja, iniciando rollback', { resultado: resultadoCaja });
        if (invSnapshot) {
          restoreSnapshot_(invSnapshot);
          logInfo('Rollback de Inventario completado');
        } else {
          logError('No se pudo hacer rollback, snapshot de Inventario no disponible');
        }
        return resultadoCaja;
      }

      logInfo('moverDeRepoACajas completado exitosamente');
      return true;
      
    } catch (e) {
      // Rollback en caso de excepción
      logError('Excepción en moverDeRepoACajas, iniciando rollback', { error: e.message, stack: e.stack });
      if (invSnapshot) restoreSnapshot_(invSnapshot);
      return `❌ Error crítico: ${e.message}`;
    }
  }, 'moverDeRepoACajas');
}

function moverDeCajasAConsumo(invSheet, codigo, lote, cantidad, cajaOrigen, filaMovimiento) {
  // Validaciones básicas
  if (!invSheet) throw new Error('Se requiere la hoja Inventario.');
  if (!codigo || !lote || !cajaOrigen || isNaN(cantidad) || cantidad <= 0) {
    return "❌ Datos inválidos. Revisa código, lote, caja origen y cantidad.";
  }

  // Restar de la caja origen (el consumo es una resta simple)
  const datosLog = { tipoMovimiento: 'Consumo', filaMovimiento };
  const resultado = restarEnInventario_(invSheet, codigo, lote, cajaOrigen, cantidad, datosLog);
  return resultado; // true si ok, mensaje de error si falla
}


function moverDeRepoADistribucion(invSheet, codigo, lote, cantidad, filaMovimiento) {
  // Validaciones básicas
  if (!invSheet) throw new Error('Se requiere la hoja Inventario.');
  if (!codigo || !lote || isNaN(cantidad) || cantidad <= 0) {
    return "❌ Datos inválidos. Revisa código, lote y cantidad.";
  }

  // Restar de Depo (la distribución es una resta simple)
  const datosLog = { tipoMovimiento: 'Distribución', filaMovimiento };
  const resultado = restarEnInventario_(invSheet, codigo, lote, 'Depo', cantidad, datosLog);
  return resultado; // true si ok, mensaje de error si falla
}


