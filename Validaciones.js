/* function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Validaciones')
    .addItem('🔍 Validar lotes desde celda...', 'pedirInicioColumna')
    .addToUi();
} */
/************  MENÚ PERSONALIZADO ************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🧰 Gestión de Stock")
    .addItem("📦 Registrar Movimiento (Formulario QR)", "mostrarFormularioQR")
    .addItem("🧾 Ajuste de Stock", "mostrarFormularioAjusteStock")
    .addSeparator()
    .addItem("↩️ Deshacer Movimiento...", "mostrarPromptDeshacerMovimiento")
    .addSeparator()
    .addItem("🔄 Actualizar Stock Total", "actualizarStockTotal")
    .addSeparator()
    .addItem("📋 Generar Trazabilidad Mensual", "generarTrazabilidadMes")
    .addSeparator()
    .addItem("↩️ Procesar Re-ingresos", "procesarReIngresos")
    .addSeparator()
    .addSubMenu(ui.createMenu("🔧 Utilidades")
      .addItem('🔍 Validar lotes desde celda...', 'pedirInicioColumna')
      .addItem('🔐 Validar IDs de Inventario', 'validarIDsInventario')
      .addSeparator()
      .addItem('⚡ Migrar Inventario (generar IDs)', 'migrarInventarioConIDs')
      .addItem('📋 Verificar encabezados Inventario', 'verificarEncabezadosManual')
      .addItem('🧹 Limpiar logs largos', 'limpiarLogsLargos')
      .addItem('🗑️ Limpiar filas en cero (opcional)', 'limpiarCerosManual'))
    .addToUi();
}

// Verifica y corrige encabezados de Inventario
function verificarEncabezadosManual() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName('Inventario');
  
  if (!invSheet) {
    ui.alert('❌ No se encontró la hoja Inventario');
    return;
  }
  
  // Leer encabezados actuales
  const primeraFila = invSheet.getRange(1, 1, 1, 6).getValues()[0];
  let mensaje = '📋 Encabezados actuales:\n\n';
  mensaje += `A: "${primeraFila[0]}"\n`;
  mensaje += `B: "${primeraFila[1]}"\n`;
  mensaje += `C: "${primeraFila[2]}"\n`;
  mensaje += `D: "${primeraFila[3]}"\n`;
  mensaje += `E: "${primeraFila[4]}"\n`;
  mensaje += `F: "${primeraFila[5]}"\n\n`;
  mensaje += '¿Quieres corregirlos al formato esperado?';
  
  const respuesta = ui.alert('Verificar Encabezados', mensaje, ui.ButtonSet.YES_NO);
  
  if (respuesta === ui.Button.YES) {
    verificarYCrearEncabezados_(invSheet);
    
    if (typeof invalidarCacheColumnas === 'function') {
      invalidarCacheColumnas();
    }
    
    ui.alert('✅ Encabezados verificados', 
      'Los encabezados de la hoja Inventario fueron verificados y corregidos:\n\n' +
      'Columna A: ID\n' +
      'Columna B: Codigo\n' +
      'Columna C: Lote\n' +
      'Columna D: Cantidad\n' +
      'Columna E: Ubicacion\n' +
      'Columna F: Logs\n\n' +
      '⚠️ IMPORTANTE: Cierra y vuelve a abrir el archivo para que los cambios surtan efecto completo.',
      ui.ButtonSet.OK);
  }
}

// Wrapper para ejecutar limpiarCeros manualmente desde menú
function limpiarCerosManual() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Limpiar filas en cero',
    '⚠️ Esto eliminará TODAS las filas con cantidad=0 en Inventario.\n\n' +
    'Se perderá el historial de logs de esos lotes.\n\n' +
    '¿Está seguro de continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (respuesta === ui.Button.YES) {
    limpiarCeros(['Inventario']);
    ui.alert('✅ Limpieza completada', 'Se eliminaron todas las filas con cantidad=0.', ui.ButtonSet.OK);
  }
}

function pedirInicioColumna() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.prompt(
    'Validación de lotes',
    'Ingrese la celda inicial de la columna a validar (ej: C2):',
    ui.ButtonSet.OK_CANCEL
  );

  if (respuesta.getSelectedButton() === ui.Button.OK) {
    const celdaInicial = respuesta.getResponseText().trim().toUpperCase();
    if (!/^[A-Z]+\d+$/.test(celdaInicial)) {
      ui.alert('⚠️ Formato inválido. Use por ejemplo "C2".');
      return;
    }

    validarLotesDesdeCelda(celdaInicial);
  }
}

function validarLotesDesdeCelda(celdaRef) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const celda = hoja.getRange(celdaRef);
  const columna = celda.getColumn();
  const filaInicio = celda.getRow();
  const ultimaFila = hoja.getLastRow();
  const numFilas = ultimaFila - filaInicio + 1;

  const rango = hoja.getRange(filaInicio, columna, numFilas, 1);
  const valores = rango.getValues();

  const regex = /^[A-Z0-9]+-\d{15}$/;
  const filasInvalidas = [];

  // Limpiar color anterior
  rango.setBackground(null);

  for (let i = 0; i < valores.length; i++) {
    const lote = valores[i][0];
    if (lote && typeof lote === 'string' && !regex.test(lote)) {
      filasInvalidas.push(i);
    }
  }

  // Marcar en amarillo las celdas incorrectas
  filasInvalidas.forEach(i => {
    rango.getCell(i + 1, 1).setBackground('#FFF176'); // Amarillo claro
  });

  const ui = SpreadsheetApp.getUi();
  if (filasInvalidas.length > 0) {
    ui.alert(`✅ Validación completada. Se encontraron ${filasInvalidas.length} lote(s) inválido(s).`);
  } else {
    ui.alert('🎉 Todos los lotes son válidos.');
  }
}
