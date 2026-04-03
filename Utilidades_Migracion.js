/**
 * SCRIPT DE MIGRACIÓN - Ejecutar UNA SOLA VEZ
 * Genera IDs y logs iniciales para registros existentes en Inventario
 */

// ─── Configuración de entorno ────────────────────────────────────────────────
// Correr UNA SOLA VEZ en cada proyecto GAS (dev y prod) desde el editor de Apps Script.
// Menú superior → Ejecutar → configurarEntornoDev / configurarEntornoProd

function configurarEntornoDev() {
  PropertiesService.getScriptProperties().setProperty(
    'ID_AGENDA', '1AGvxp31Wwe6nM-tiT-I_hpg1S49aMtxEJY_8kkP-MEoVC3r-adQJfv-L'
  );
  SpreadsheetApp.getUi().alert('✅ Entorno DEV configurado correctamente.');
}

function configurarEntornoProd() {
  PropertiesService.getScriptProperties().setProperty(
    'ID_AGENDA', '1eAzSrs1AFKljA8VY_3vDVxBSjEGXml8rBMr2SgNJrEbEBmJyurMPj0IF'
  );
  SpreadsheetApp.getUi().alert('✅ Entorno PROD configurado correctamente.');
}

function migrarInventarioConIDs() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    '⚠️ Migración de Inventario',
    '¿Estás seguro de que querés generar IDs y logs para todos los registros existentes?\n\n' +
    'Esta operación:\n' +
    '• Generará IDs únicos para cada código+lote\n' +
    '• Agregará log inicial con timestamp actual\n' +
    '• NO modificará cantidades ni ubicaciones\n' +
    '• Puede tardar varios segundos\n\n' +
    '⚠️ IMPORTANTE: Ejecutar solo UNA vez',
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    ui.alert('✅ Migración cancelada');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = ss.getSheetByName('Inventario');
    
    if (!invSheet) {
      ui.alert('❌ Error: No se encontró la hoja Inventario');
      return;
    }

    // ✅ PASO 1: Verificar y crear encabezados si no existen
    verificarYCrearEncabezados_(invSheet);
    
    // ✅ PASO 2: Invalidar cache de columnas para forzar re-lectura
    if (typeof invalidarCacheColumnas === 'function') {
      invalidarCacheColumnas();
    }

    // Usar índices por defecto (hardcoded) para migración inicial
    // Ya que las columnas ID y Logs pueden estar vacías sin encabezado
    const COL_ID = getDefaultColumnIndexInventario('id');       // Col A = 1
    const COL_COD = getDefaultColumnIndexInventario('codigo');  // Col B = 2
    const COL_LOTE = getDefaultColumnIndexInventario('lote');   // Col C = 3
    const COL_CANT = getDefaultColumnIndexInventario('cantidad'); // Col D = 4
    const COL_UBIC = getDefaultColumnIndexInventario('ubicacion'); // Col E = 5
    const COL_LOGS = getDefaultColumnIndexInventario('logs');   // Col F = 6

    const lastRow = invSheet.getLastRow();
    const lastCol = invSheet.getLastColumn();
    
    if (lastRow < 2) {
      ui.alert('ℹ️ Info: No hay datos para migrar en Inventario');
      return;
    }

    const data = invSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    let actualizados = 0;
    let omitidos = 0;
    
    const fecha = new Date();
    const tz = Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires';
    const fechaStr = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd HH:mm:ss');

    for (let i = 0; i < data.length; i++) {
      const fila = i + 2;
      const id = (data[i][COL_ID - 1] || '').toString().trim();
      const codigo = (data[i][COL_COD - 1] || '').toString().trim().toUpperCase();
      const lote = (data[i][COL_LOTE - 1] || '').toString().trim().toUpperCase();
      const cantidad = Number(data[i][COL_CANT - 1] || 0);
      const ubicacion = (data[i][COL_UBIC - 1] || '').toString().trim();
      const logs = (data[i][COL_LOGS - 1] || '').toString().trim();

      // Omitir filas vacías o sin cantidad
      if (!codigo || !lote || cantidad <= 0) {
        omitidos++;
        continue;
      }

      let necesitaActualizacion = false;

      // Generar ID si no existe
      if (!id) {
        const nuevoID = generarIDInventario_(codigo, lote, ubicacion);
        invSheet.getRange(fila, COL_ID).setValue(nuevoID);
        necesitaActualizacion = true;
      }

      // Generar log inicial si no existe
      if (!logs) {
        const logInicial = `${fechaStr} | Cantidad inicial: ${cantidad} en ${ubicacion} (migración)`;
        invSheet.getRange(fila, COL_LOGS).setValue(logInicial);
        necesitaActualizacion = true;
      }

      if (necesitaActualizacion) {
        actualizados++;
      } else {
        omitidos++;
      }
    }

    SpreadsheetApp.flush();
    
    // ✅ PASO 3: Invalidar cache nuevamente después de migración
    if (typeof invalidarCacheColumnas === 'function') {
      invalidarCacheColumnas();
      Logger.log('Cache de columnas invalidado después de migración');
    }
    
    ui.alert(
      '✅ Migración Completada',
      `Se procesaron ${lastRow - 1} filas:\n\n` +
      `• Actualizadas: ${actualizados}\n` +
      `• Omitidas (ya tenían ID/logs o sin datos): ${omitidos}\n\n` +
      `Fecha de migración: ${fechaStr}`,
      ui.ButtonSet.OK
    );

    // Ejecutar validación de integridad
    const validacion = validarIntegridadInventario_();
    if (!validacion.valido) {
      ui.alert(
        '⚠️ Advertencia',
        `Se encontraron ${validacion.errores.length} errores de integridad.\n` +
        `Revisá el log de ejecución para más detalles.`,
        ui.ButtonSet.OK
      );
      Logger.log('Errores de validación post-migración:');
      validacion.errores.forEach(err => Logger.log(`  - ${err}`));
    }

  } catch (error) {
    ui.alert('❌ Error durante la migración: ' + error.toString());
    Logger.log('Error en migración: ' + error.toString());
    Logger.log(error.stack);
  }
}

/**
 * Función para validar que no haya IDs duplicados en la misma ubicación
 * Ejecutar después de migraciones o cambios manuales
 */
function validarIDsInventario() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const validacion = validarIntegridadInventario_();
    
    let mensaje = `📊 Validación de Inventario\n\n`;
    mensaje += `Total de filas: ${validacion.totalFilas}\n`;
    mensaje += `Errores: ${validacion.errores.length}\n`;
    mensaje += `Advertencias: ${validacion.advertencias.length}\n\n`;
    
    if (validacion.valido) {
      mensaje += '✅ No se encontraron errores críticos';
    } else {
      mensaje += '❌ SE ENCONTRARON ERRORES:\n\n';
      validacion.errores.slice(0, 5).forEach(err => {
        mensaje += `• ${err}\n`;
      });
      if (validacion.errores.length > 5) {
        mensaje += `\n... y ${validacion.errores.length - 5} errores más (ver log)`;
      }
    }
    
    if (validacion.advertencias.length > 0) {
      mensaje += '\n\n⚠️ ADVERTENCIAS:\n';
      validacion.advertencias.slice(0, 3).forEach(adv => {
        mensaje += `• ${adv}\n`;
      });
      if (validacion.advertencias.length > 3) {
        mensaje += `\n... y ${validacion.advertencias.length - 3} advertencias más (ver log)`;
      }
    }
    
    ui.alert('Validación de Inventario', mensaje, ui.ButtonSet.OK);
    
    // Log completo
    Logger.log('=== VALIDACIÓN DE INVENTARIO ===');
    Logger.log(`Total filas: ${validacion.totalFilas}`);
    Logger.log(`Válido: ${validacion.valido}`);
    
    if (validacion.errores.length > 0) {
      Logger.log('\nERRORES:');
      validacion.errores.forEach(err => Logger.log(`  ${err}`));
    }
    
    if (validacion.advertencias.length > 0) {
      Logger.log('\nADVERTENCIAS:');
      validacion.advertencias.forEach(adv => Logger.log(`  ${adv}`));
    }
    
  } catch (error) {
    ui.alert('❌ Error al validar: ' + error.toString());
    Logger.log('Error en validación: ' + error.toString());
  }
}

/**
 * Función para limpiar logs muy largos (más de 5000 caracteres)
 * Mantiene solo los últimos 10 movimientos
 */
function limpiarLogsLargos() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Limpiar Logs Largos',
    '¿Querés limpiar los logs que superen 5000 caracteres?\n\n' +
    'Se mantendrán solo los últimos 10 movimientos de cada log.',
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    ui.alert('✅ Operación cancelada');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = ss.getSheetByName('Inventario');
    const COL_LOGS = getDefaultColumnIndexInventario('logs'); // Usar índice fijo (col 6)
    
    const lastRow = invSheet.getLastRow();
    if (lastRow < 2) return;
    
    const logs = invSheet.getRange(2, COL_LOGS, lastRow - 1, 1).getValues();
    let limpiados = 0;
    
    for (let i = 0; i < logs.length; i++) {
      const log = (logs[i][0] || '').toString();
      
      if (log.length > 5000) {
        const movimientos = log.split(';').map(m => m.trim());
        const ultimos10 = movimientos.slice(-10);
        const nuevoLog = ultimos10.join('; ');
        
        invSheet.getRange(i + 2, COL_LOGS).setValue(nuevoLog);
        limpiados++;
      }
    }
    
    SpreadsheetApp.flush();
    ui.alert(`✅ Limpieza completada\n\nLogs limpiados: ${limpiados}`);
    
  } catch (error) {
    ui.alert('❌ Error en limpieza: ' + error.toString());
    Logger.log('Error en limpieza de logs: ' + error.toString());
  }
}

/**
 * Verifica y crea encabezados en la hoja Inventario si no existen
 * @param {Sheet} invSheet - Hoja de Inventario
 */
function verificarYCrearEncabezados_(invSheet) {
  const primeraFila = invSheet.getRange(1, 1, 1, 6).getValues()[0];
  const encabezadosEsperados = ['ID', 'Codigo', 'Lote', 'Cantidad', 'Ubicacion', 'Logs'];
  
  let necesitaActualizar = false;
  
  for (let i = 0; i < encabezadosEsperados.length; i++) {
    const valorActual = (primeraFila[i] || '').toString().trim();
    const esperado = encabezadosEsperados[i];
    
    // Si está vacío o no coincide (case-insensitive), actualizar
    if (!valorActual || valorActual.toLowerCase() !== esperado.toLowerCase()) {
      invSheet.getRange(1, i + 1).setValue(esperado);
      necesitaActualizar = true;
      Logger.log(`Encabezado creado/actualizado en columna ${i + 1}: "${esperado}"`);
    }
  }
  
  if (necesitaActualizar) {
    SpreadsheetApp.flush();
    Logger.log('Encabezados de Inventario verificados y actualizados');
  }
}
