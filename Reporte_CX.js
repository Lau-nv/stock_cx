// ─── Configuración ───────────────────────────────────────────────────────────
// ID del Google Sheet de la Agenda CX (cambiar aquí para producción)
const ID_AGENDA = '1Kg3J6dTS2SaUvz5AhDB8i6hgfLrUO_E9j43ymNxrxoM';

// Columnas de Agenda Cx (0-based desde columna A)
const AGENDA_COL_FECHA       = 0; // A
const AGENDA_COL_ID          = 2; // C
const AGENDA_COL_PACIENTE    = 3; // D
const AGENDA_COL_INSTITUCION = 4; // E
const AGENDA_COL_MEDICO      = 6; // G
const AGENDA_COL_CLIENTE     = 7; // H

const MESES_ES = [
  'Enero','Febrero','Marzo','Abril','Mayo','Junio',
  'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
];

const TRAZABILIDAD_HEADERS = [
  'Tipo', 'ID CX', 'Fecha CX', 'Paciente', 'Institución', 'Cliente', 'Médico',
  'Código', 'Lote', 'Cantidad', 'Fecha Movimiento'
];

// Tipos que deben excluirse siempre del reporte
const TIPOS_EXCLUIDOS = new Set([
  'reposicion', 'reposicion caja completa', 'entre cajas',
  'ingreso', 'ingreso desde liberaciones'
]);

// ─── Puntos de entrada ───────────────────────────────────────────────────────

// Abre el diálogo HTML con selectores de mes y año
function generarTrazabilidadMes() {
  const html = HtmlService.createHtmlOutputFromFile('Dialogo_Trazabilidad')
    .setWidth(320)
    .setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generar Trazabilidad');
}

// Llamado desde el diálogo HTML. mes = 0-based (0=Enero, 11=Diciembre)
function generarTrazabilidadDesdeMes(mes, ano) {
  const movSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Movimientos');
  if (!movSheet || movSheet.getLastRow() < 2) return '❌ No hay movimientos registrados.';
  const agendaMap  = obtenerMapaAgenda_();
  const filas      = construirFilasMes_(mes, ano, movSheet, agendaMap);
  const nombreHoja = `Trazabilidad ${MESES_ES[mes]} ${ano}`;
  escribirHoja_(nombreHoja, filas);
  return `✅ "${nombreHoja}" generada con ${filas.length} registro/s.`;
}

// Regenera todas las hojas para todos los meses con datos
function regenerarTrazabilidadCompleta() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const movSheet = ss.getSheetByName('Movimientos');
  if (!movSheet || movSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No hay movimientos registrados.');
    return;
  }
  const agendaMap = obtenerMapaAgenda_();
  const meses     = obtenerMesesRelevantes_(movSheet, agendaMap);
  if (meses.length === 0) {
    SpreadsheetApp.getUi().alert('No hay movimientos para incluir en la Trazabilidad.');
    return;
  }
  meses.forEach(({ mes, ano }) => {
    const filas      = construirFilasMes_(mes, ano, movSheet, agendaMap);
    const nombreHoja = `Trazabilidad ${MESES_ES[mes]} ${ano}`;
    escribirHoja_(nombreHoja, filas);
  });
  const etiquetas = meses.map(({ mes, ano }) => `${MESES_ES[mes]} ${ano}`).join(', ');
  SpreadsheetApp.getUi().alert(`✅ Trazabilidad regenerada para: ${etiquetas}`);
}

// ─── Lógica interna ──────────────────────────────────────────────────────────

function construirFilasMes_(mes, ano, movSheet, agendaMap) {
  const lastRow  = movSheet.getLastRow();
  if (lastRow < 2) return [];

  const movData  = movSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const anuladas = getFilasAnuladas_(movData);
  const filas    = [];

  movData.forEach((row, i) => {
    const sheetRow = i + 2; // número real de fila en el sheet (fila 1 = header)

    // 1. Saltar movimientos anulados
    if (anuladas.has(sheetRow)) return;

    const tipoRaw = (row[1] || '').toString();
    const tipo    = normalizarTipo_(tipoRaw);

    // 2. Saltar anulaciones y tipos no relevantes
    if (tipo.startsWith('anulaci') || TIPOS_EXCLUIDOS.has(tipo)) return;

    const idCx    = (row[10] || '').toString().trim();
    const fechaMov = row[0] ? new Date(row[0]) : null;

    if (tipo === 'consumo') {
      // Fecha de referencia: Fecha CX si existe, sino Fecha Movimiento
      const cx       = agendaMap[idCx] || {};
      const fechaCX  = cx.fecha ? new Date(cx.fecha) : null;
      const fechaRef = fechaCX || fechaMov;
      if (!fechaRef || fechaRef.getMonth() !== mes || fechaRef.getFullYear() !== ano) return;
      filas.push([
        tipoRaw,
        idCx           || 'N/A',
        cx.fecha       || 'N/A',
        cx.paciente    || 'N/A',
        cx.institucion || 'N/A',
        cx.cliente     || 'N/A',
        cx.medico      || 'N/A',
        row[2], row[3], row[4], row[0]
      ]);

    } else if (tipo === 'egreso' || tipo === 'distribucion') {
      if (!fechaMov || fechaMov.getMonth() !== mes || fechaMov.getFullYear() !== ano) return;
      filas.push([
        tipoRaw,
        'N/A', 'N/A', 'N/A', 'N/A',
        (row[8] || 'N/A').toString(), // Cliente desde Movimientos
        'N/A',
        row[2], row[3], row[4], row[0]
      ]);
    }
  });

  return filas;
}

function escribirHoja_(nombreHoja, filas) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  let hoja   = ss.getSheetByName(nombreHoja);
  if (!hoja) hoja = ss.insertSheet(nombreHoja);
  hoja.clearContents();
  hoja.getRange(1, 1, 1, TRAZABILIDAD_HEADERS.length).setValues([TRAZABILIDAD_HEADERS]);
  if (filas.length > 0) {
    hoja.getRange(2, 1, filas.length, TRAZABILIDAD_HEADERS.length).setValues(filas);
  }
  logInfo('Trazabilidad escrita', { hoja: nombreHoja, filas: filas.length });
}

// Detecta todos los meses/años con movimientos relevantes
function obtenerMesesRelevantes_(movSheet, agendaMap) {
  const lastRow  = movSheet.getLastRow();
  if (lastRow < 2) return [];
  const movData  = movSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const anuladas = getFilasAnuladas_(movData);
  const claves   = new Set();

  movData.forEach((row, i) => {
    if (anuladas.has(i + 2)) return;
    const tipo  = normalizarTipo_(row[1]);
    if (tipo.startsWith('anulaci') || TIPOS_EXCLUIDOS.has(tipo)) return;

    const idCx    = (row[10] || '').toString().trim();
    const fechaMov = row[0] ? new Date(row[0]) : null;

    if (tipo === 'consumo') {
      const cx      = agendaMap[idCx] || {};
      const fechaCX = cx.fecha ? new Date(cx.fecha) : null;
      const fechaRef = fechaCX || fechaMov;
      if (fechaRef && !isNaN(fechaRef)) {
        claves.add(`${fechaRef.getFullYear()}-${fechaRef.getMonth()}`);
      }
    } else if (tipo === 'egreso' || tipo === 'distribucion') {
      if (fechaMov && !isNaN(fechaMov)) {
        claves.add(`${fechaMov.getFullYear()}-${fechaMov.getMonth()}`);
      }
    }
  });

  return [...claves]
    .map(k => { const [ano, mes] = k.split('-').map(Number); return { ano, mes }; })
    .sort((a, b) => a.ano !== b.ano ? a.ano - b.ano : a.mes - b.mes);
}

// Construye un Set con los números de fila (1-based en sheet) que fueron anulados.
// Lee las filas de tipo "Anulación X" y extrae el número de fila original
// del texto de observaciones: "ANULACIÓN de fila N (tipo: ...)"
function getFilasAnuladas_(movData) {
  const anuladas = new Set();
  movData.forEach(row => {
    if (!normalizarTipo_(row[1]).startsWith('anulaci')) return;
    const match = (row[9] || '').toString().match(/de fila (\d+)/i);
    if (match) anuladas.add(parseInt(match[1]));
  });
  return anuladas;
}

// Devuelve mapa { idCx -> { fecha, paciente, institucion, cliente, medico } }
// Lee todas las hojas cuyo nombre empiece con "Agenda"
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
      // Lee columnas A-J (10 columnas); J=índice 9 = Estado
      sheet.getRange(2, 1, lastRow - 1, 10).getValues().forEach(row => {
        const id     = (row[AGENDA_COL_ID] || '').toString().trim();
        const estado = (row[9]             || '').toString().trim();
        if (!id || estado !== 'Autorizada') return;
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
