// Columnas de la hoja "Re-ingresos" (1-based para getRange, -1 para acceso a array)
const RI_COL = {
  CHECK:          1,  // A - Checkbox nativo
  FECHA_MOV:      2,  // B - Fecha del movimiento original
  TIPO:           3,  // C - Tipo (Egreso / Distribución / etc.)
  CODIGO:         4,  // D - Código
  LOTE:           5,  // E - Lote
  CANTIDAD:       6,  // F - Cantidad
  DESTINO:        7,  // G - Destino / Cliente
  OBS:            8,  // H - Observaciones originales
  ESTADO:         9,  // I - "" = pendiente | "Devuelto" = procesado
  FECHA_REINGRESO:10  // J - Fecha en que se procesó el re-ingreso
};

const NOMBRE_HOJA_REINGRESOS = 'Re-ingresos';

/* ─── Creación / inicialización de la hoja ──────────────────────────────── */
function crearHojaReingresos_(ss) {
  let hoja = ss.getSheetByName(NOMBRE_HOJA_REINGRESOS);
  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_HOJA_REINGRESOS);
  }

  const headers = [
    'Re-ingresar?', 'Fecha Movimiento', 'Tipo', 'Código', 'Lote',
    'Cantidad', 'Destino/Cliente', 'Observaciones', 'Estado', 'Fecha Re-ingreso'
  ];
  const headerRange = hoja.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4a4a4a')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  hoja.setFrozenRows(1);

  const anchos = [90, 140, 100, 80, 150, 70, 150, 220, 90, 140];
  anchos.forEach((w, i) => hoja.setColumnWidth(i + 1, w));

  return hoja;
}

/* ─── Agregar una fila a Re-ingresos ────────────────────────────────────── */
/**
 * Se llama después de registrar un egreso o distribución exitoso.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Date}   fecha        - Fecha del movimiento original
 * @param {string} tipo         - Tipo legible (ej: "Egreso", "Distribución")
 * @param {string} codigo
 * @param {string} lote
 * @param {number} cantidad
 * @param {string} clienteDestino
 * @param {string} observaciones
 */
function agregarAReIngresos_(ss, fecha, tipo, codigo, lote, cantidad, clienteDestino, observaciones) {
  let hoja = ss.getSheetByName(NOMBRE_HOJA_REINGRESOS);
  if (!hoja) hoja = crearHojaReingresos_(ss);

  const nextRow = hoja.getLastRow() + 1;

  hoja.appendRow([
    false,                          // A: checkbox (se reemplaza abajo)
    fecha || new Date(),            // B: fecha movimiento original
    tipo   || 'N/A',                // C: tipo
    codigo || '',                   // D: código
    lote   || '',                   // E: lote
    cantidad,                       // F: cantidad
    clienteDestino || 'N/A',        // G: destino/cliente
    observaciones  || 'N/A',        // H: observaciones
    '',                             // I: estado (vacío = pendiente)
    ''                              // J: fecha re-ingreso
  ]);

  // Reemplazar el FALSE por un checkbox nativo de Sheets
  hoja.getRange(nextRow, RI_COL.CHECK).insertCheckboxes();
}

/* ─── Procesar ítems tildados ───────────────────────────────────────────── */
function procesarReIngresos() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(NOMBRE_HOJA_REINGRESOS);

  if (!hoja) {
    ui.alert('❌ No existe la hoja "Re-ingresos".');
    return;
  }

  const lastRow = hoja.getLastRow();
  if (lastRow < 2) {
    ui.alert('ℹ️ No hay ítems en la hoja Re-ingresos.');
    return;
  }

  const numCols = Object.keys(RI_COL).length; // 10
  const data = hoja.getRange(2, 1, lastRow - 1, numCols).getValues();

  const aIngresar = [];
  for (let i = 0; i < data.length; i++) {
    const check  = data[i][RI_COL.CHECK - 1];
    const estado = (data[i][RI_COL.ESTADO - 1] || '').toString().trim();
    if (check === true && estado !== 'Re-ingresado') {
      aIngresar.push({
        rowIndex: i + 2,
        tipo:     (data[i][RI_COL.TIPO     - 1] || '').toString().trim(),
        codigo:   (data[i][RI_COL.CODIGO   - 1] || '').toString().trim().toUpperCase(),
        lote:     (data[i][RI_COL.LOTE     - 1] || '').toString().trim().toUpperCase(),
        cantidad: Number(data[i][RI_COL.CANTIDAD - 1] || 0),
        obs:      (data[i][RI_COL.OBS      - 1] || '').toString().trim()
      });
    }
  }

  if (aIngresar.length === 0) {
    ui.alert('ℹ️ No hay ítems seleccionados (o todos ya fueron procesados).');
    return;
  }

  const confirmacion = ui.alert(
    '↩️ Confirmar Re-ingreso',
    `¿Confirmar el re-ingreso de ${aIngresar.length} ítem/s al inventario (Depo)?`,
    ui.ButtonSet.YES_NO
  );
  if (confirmacion !== ui.Button.YES) return;

  const invSheet = ss.getSheetByName('Inventario');
  const movSheet = ss.getSheetByName('Movimientos');
  if (!invSheet) { ui.alert('❌ No existe la hoja "Inventario".'); return; }
  if (!movSheet)  { ui.alert('❌ No existe la hoja "Movimientos".'); return; }

  const ahora = new Date();
  let ok = 0;
  const errores = [];

  for (const item of aIngresar) {
    if (!item.codigo || !item.lote || !(item.cantidad > 0)) {
      errores.push(`Fila ${item.rowIndex}: datos incompletos.`);
      continue;
    }

    const obsReingreso = `Re-ingreso (${item.tipo})` +
      (item.obs && item.obs !== 'N/A' ? ` — ${item.obs}` : '');

    const datosLog = { tipoMovimiento: 'Ingreso', observaciones: obsReingreso };
    const resultado = sumarEnInventario_(invSheet, item.codigo, item.lote, 'Depo', item.cantidad, datosLog);

    if (resultado !== true) {
      errores.push(`Fila ${item.rowIndex} (${item.codigo} / ${item.lote}): ${resultado}`);
      continue;
    }

    // Registrar en Movimientos
    movSheet.appendRow([
      ahora, 'Ingreso', item.codigo, item.lote, item.cantidad,
      'N/A', 'N/A', 'N/A', 'N/A', obsReingreso, 'N/A'
    ]);

    // Marcar fila como procesada
    hoja.getRange(item.rowIndex, RI_COL.CHECK).setValue(false);
    hoja.getRange(item.rowIndex, RI_COL.ESTADO).setValue('Re-ingresado');
    hoja.getRange(item.rowIndex, RI_COL.FECHA_REINGRESO).setValue(ahora);

    ok++;
  }

  SpreadsheetApp.flush();
  if (typeof actualizarStockTotal === 'function') actualizarStockTotal();

  let msg = `✅ Re-ingresados ${ok} ítem/s a Inventario (Depo).`;
  if (errores.length > 0) msg += `\n\n⚠️ Errores:\n${errores.join('\n')}`;
  ui.alert('Re-ingresos', msg, ui.ButtonSet.OK);
}
