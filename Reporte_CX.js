const ID_AGENDA = '1Kg3J6dTS2SaUvz5AhDB8i6hgfLrUO_E9j43ymNxrxoM';

// Columnas de Agenda Cx (0-based desde columna A)
const AGENDA_COL_FECHA       = 0; // A
const AGENDA_COL_ID          = 2; // C
const AGENDA_COL_PACIENTE    = 3; // D
const AGENDA_COL_INSTITUCION = 4; // E
const AGENDA_COL_MEDICO      = 6; // G
const AGENDA_COL_CLIENTE     = 7; // H

const TRAZABILIDAD_HEADERS = [
  'ID CX', 'Fecha CX', 'Paciente', 'Institución', 'Cliente', 'Médico',
  'Código', 'Lote', 'Cantidad', 'Fecha Movimiento'
];

// Llamado automáticamente después de cada Consumo exitoso.
// Recibe el idCx directamente para determinar el año de la CX sin ambigüedad.
function generarReporteCX(idCx) {
  if (!idCx || idCx === 'N/A') return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const movSheet = ss.getSheetByName('Movimientos');
  if (!movSheet || movSheet.getLastRow() < 2) return;

  const agendaMap = obtenerMapaAgenda_();
  const cx = agendaMap[idCx];
  if (!cx || !cx.fecha) {
    logWarn('generarReporteCX: ID CX no encontrado en agenda', { idCx });
    return;
  }

  const yearCX = new Date(cx.fecha).getFullYear();
  if (isNaN(yearCX)) {
    logWarn('generarReporteCX: fecha CX inválida', { idCx, fecha: cx.fecha });
    return;
  }

  generarTrazabilidadAno_(yearCX, movSheet, agendaMap);
}

// Regenera todas las hojas de Trazabilidad (una por año de Fecha CX).
// Para uso manual desde el menú.
function regenerarTrazabilidadCompleta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const movSheet = ss.getSheetByName('Movimientos');
  if (!movSheet || movSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No hay movimientos registrados.');
    return;
  }

  const agendaMap = obtenerMapaAgenda_();
  const anos = obtenerAnosDesdeCX_(movSheet, agendaMap);

  if (anos.size === 0) {
    SpreadsheetApp.getUi().alert('No hay Consumos con ID CX vinculado a la agenda.');
    return;
  }

  anos.forEach(year => generarTrazabilidadAno_(year, movSheet, agendaMap));
  SpreadsheetApp.getUi().alert(`✅ Trazabilidad regenerada para: ${[...anos].sort().join(', ')}`);
}

// ─── Interno ────────────────────────────────────────────────────────────────

// Retorna un Set con los años (de Fecha CX) presentes en los Consumos con ID CX válido
function obtenerAnosDesdeCX_(movSheet, agendaMap) {
  const lastRow = movSheet.getLastRow();
  if (lastRow < 2) return new Set();
  const movData = movSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const anos = new Set();
  movData.forEach(row => {
    const tipo = normalizarTipo_(row[1]);
    const idCx = (row[10] || '').toString().trim();
    if (tipo !== 'consumo' || !idCx || idCx === 'N/A') return;
    const cx = agendaMap[idCx];
    if (!cx || !cx.fecha) return;
    const y = new Date(cx.fecha).getFullYear();
    if (!isNaN(y)) anos.add(y);
  });
  return anos;
}

// Construye o reconstruye la hoja "Trazabilidad YYYY" filtrando por año de Fecha CX
function generarTrazabilidadAno_(year, movSheet, agendaMap) {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHoja = 'Trazabilidad ' + year;

  const lastRow = movSheet.getLastRow();
  const consumos = lastRow < 2 ? [] : movSheet.getRange(2, 1, lastRow - 1, 11).getValues()
    .filter(row => {
      const tipo = normalizarTipo_(row[1]);
      const idCx = (row[10] || '').toString().trim();
      if (tipo !== 'consumo' || !idCx || idCx === 'N/A') return false;
      const cx = agendaMap[idCx];
      if (!cx || !cx.fecha) return false;
      return new Date(cx.fecha).getFullYear() === year;
    });

  let hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) hoja = ss.insertSheet(nombreHoja);
  hoja.clearContents();
  hoja.getRange(1, 1, 1, TRAZABILIDAD_HEADERS.length).setValues([TRAZABILIDAD_HEADERS]);

  if (consumos.length === 0) return;

  const filas = consumos.map(row => {
    const idCx = (row[10] || '').toString().trim();
    const cx   = agendaMap[idCx] || {};
    return [
      idCx,
      cx.fecha        || 'N/A',
      cx.paciente     || 'N/A',
      cx.institucion  || 'N/A',
      cx.cliente      || 'N/A',
      cx.medico       || 'N/A',
      row[2],  // Código
      row[3],  // Lote
      row[4],  // Cantidad
      row[0]   // Fecha Movimiento
    ];
  });

  hoja.getRange(2, 1, filas.length, TRAZABILIDAD_HEADERS.length).setValues(filas);
  logInfo('Trazabilidad generada', { hoja: nombreHoja, filas: filas.length });
}

// Devuelve mapa { idCx -> { fecha, paciente, institucion, cliente, medico } }
// Lee todas las hojas del spreadsheet de agenda cuyo nombre empiece con "Agenda"
function obtenerMapaAgenda_() {
  try {
    const agendaSS = SpreadsheetApp.openById(ID_AGENDA);
    const hojas    = agendaSS.getSheets().filter(h =>
      h.getName().toLowerCase().startsWith('agenda')
    );
    const map = {};
    hojas.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      data.forEach(row => {
        const id = (row[AGENDA_COL_ID] || '').toString().trim();
        if (!id) return;
        map[id] = {
          fecha:       row[AGENDA_COL_FECHA],
          paciente:    (row[AGENDA_COL_PACIENTE]    || '').toString().trim(),
          institucion: (row[AGENDA_COL_INSTITUCION] || '').toString().trim(),
          medico:      (row[AGENDA_COL_MEDICO]      || '').toString().trim(),
          cliente:     (row[AGENDA_COL_CLIENTE]     || '').toString().trim()
        };
      });
    });
    return map;
  } catch (e) {
    logError('Error leyendo Agenda para trazabilidad', { error: e.message });
    return {};
  }
}
