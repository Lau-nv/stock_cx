// ─── Configuración ────────────────────────────────────────────────────────────
// ID del Google Sheet de la Agenda CX (cambiar aquí para producción)
// ID de la Agenda CX — se detecta automáticamente según el entorno.
const _SCRIPT_ID_DEV  = '1NC5je_flA1yZE2hGEbS_kCABUImJXGt9rYZ_NLVrQx2wkKV0Bp9r8SA4';
const _AGENDA_ID_DEV  = '1AGvxp31Wwe6nM-tiT-I_hpg1S49aMtxEJY_8kkP-MEoVC3r-adQJfv-L';
const _AGENDA_ID_PROD = '1eAzSrs1AFKljA8VY_3vDVxBSjEGXml8rBMr2SgNJrEbEBmJyurMPj0IF';
const ID_AGENDA = ScriptApp.getScriptId() === _SCRIPT_ID_DEV ? _AGENDA_ID_DEV : _AGENDA_ID_PROD;

// Columnas de Agenda CX (0-based desde columna A)
const AGENDA_COL_FECHA       = 0; // A
const AGENDA_COL_ID          = 2; // C
const AGENDA_COL_PACIENTE    = 3; // D
const AGENDA_COL_INSTITUCION = 4; // E
const AGENDA_COL_MEDICO      = 6;  // G
const AGENDA_COL_CLIENTE     = 7;  // H
const AGENDA_COL_FACTURA     = 15; // P

const MESES_ES = [
  'Enero','Febrero','Marzo','Abril','Mayo','Junio',
  'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
];

// Tipos que se muestran en el detalle del reporte.
// Reposición y Entre cajas se usan SOLO para calcular stock, no aparecen en el detalle.
const TIPOS_DETALLE_TRAZA = new Set([
  'ingreso', 'ingreso desde liberaciones', 'consumo', 'distribucion', 'egreso'
]);

// Resumen y detalle usan las MISMAS columnas (A-K).
// La fila resumen usa A-D; las filas detalle usan A-K.
// Esto evita tener que desplazarse horizontalmente al expandir un grupo.
const TRAZA_NUM_COLS = 12;

// Columnas de la fila resumen (0-based)
const TRAZA_S = {
  CODIGO      : 0,  // A
  LOTE        : 1,  // B
  STOCK_INI   : 2,  // C
  STOCK_FIN   : 3,  // D
  UBICACION   : 4,  // E  ← "DEPO: 6 | CC-02: 3 | CC-03: 3" (stock final)
};

// Columnas de las filas detalle (mismas posiciones, contenido diferente)
const TRAZA_D = {
  FECHA_MOV    : 0,  // A
  TIPO         : 1,  // B
  CANT         : 2,  // C
  UBICACION    : 3,  // D
  PACIENTE     : 4,  // E
  CLIENTE      : 5,  // F
  MEDICO       : 6,  // G
  INSTITUCION  : 7,  // H
  ID_CX        : 8,  // I
  FECHA_CX     : 9,  // J
  OBSERVACIONES: 10, // K
  FACTURA      : 11, // L
};

// Cabecera global (describe las filas resumen)
const TRAZA_GLOBAL_HEADER = [
  'Código', 'Lote', 'Stock Inicio', 'Stock Final', 'Ubicación', '', '', '', '', '', '', ''
];

// Mini-cabecera interna de cada grupo (describe las filas detalle)
const TRAZA_DETAIL_HEADER = [
  'Fecha', 'Tipo', 'Cantidad', 'Ubicación',
  'Paciente', 'Cliente', 'Médico', 'Institución', 'ID CX', 'Fecha CX', 'Observaciones', 'Nro. Factura'
];

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
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const shMov = ss.getSheetByName('Movimientos');
    if (!shMov || shMov.getLastRow() < 2) return '❌ No hay movimientos registrados.';

    const agendaMap  = obtenerMapaAgenda_();
    const nombreHoja = `Trazabilidad ${MESES_ES[mes]} ${ano}`;

    // Leer todos los movimientos
    const lastRow = shMov.getLastRow();
    const movData = shMov.getRange(2, 1, lastRow - 1, 11).getValues();

    // Armar lista de movimientos efectivos (excluye anulados y anulaciones)
    const filasAnuladas = getFilasAnuladas_(movData);
    const movEfectivos  = [];
    for (let i = 0; i < movData.length; i++) {
      const sheetRow = i + 2;
      const tipo     = normalizarTipo_((movData[i][1] || '').toString());
      if (filasAnuladas.has(sheetRow)) continue;
      if (tipo.startsWith('anulaci')) continue;
      movEfectivos.push({ row: movData[i], tipo });
    }

    // Límites del mes
    const dInicio  = new Date(ano, mes, 1);
    const dFinExcl = new Date(ano, mes + 1, 1); // exclusivo (= inicio del mes siguiente)

    // Encontrar pares código/lote con al menos una SALIDA en el mes
    // (consumo, egreso o distribución). Los ingresos solos no se incluyen
    // porque no aportan trazabilidad de uso del implante.
    const TIPOS_SALIDA = new Set(['consumo', 'egreso', 'distribucion']);
    const paresConMov = new Set();
    for (const { row, tipo } of movEfectivos) {
      if (!TIPOS_SALIDA.has(tipo)) continue;
      const fecha = trazaToDate_(row[0]);
      if (!fecha || fecha < dInicio || fecha >= dFinExcl) continue;
      const codigo = trazaNorm_(row[2]);
      const lote   = trazaNorm_(row[3]);
      if (codigo && lote) paresConMov.add(`${codigo}|||${lote}`);
    }

    if (paresConMov.size === 0) {
      return `⚠️ No se encontraron movimientos en ${MESES_ES[mes]} ${ano}.`;
    }

    // ── Cálculo de stock ────────────────────────────────────────────────────
    // TOTAL (por código/lote): parte del inventario actual y deshace hacia atrás.
    // Maneja correctamente consumos con cajaOrigen=N/A.
    const mapActual = leerInventarioActual_(ss);
    const movsPostMes    = movEfectivos.filter(({ row }) => { const f = trazaToDate_(row[0]); return f && f >= dFinExcl; });
    const movsDuranteMes = movEfectivos.filter(({ row }) => { const f = trazaToDate_(row[0]); return f && f >= dInicio && f < dFinExcl; });

    const mapFin = Object.assign({}, mapActual);
    deshacerMovimientos_(mapFin, movsPostMes);
    const mapInicio = Object.assign({}, mapFin);
    deshacerMovimientos_(mapInicio, movsDuranteMes);

    // DESGLOSE POR UBICACIÓN: mismo enfoque pero rastreando ubicación.
    // Si cajaOrigen='N/A' el movimiento se omite del desglose (no afecta el total).
    const mapActualUbic = leerInventarioUbicacion_(ss);
    const mapFinUbic = Object.assign({}, mapActualUbic);
    deshacerPorUbicacion_(mapFinUbic, movsPostMes);
    // mapFinUbic es suficiente: solo mostramos ubicación del stock al cierre del mes

    // ── Construir filas ──────────────────────────────────────────────────────
    // Estructura por grupo:
    //   filas[N-1]   = fila resumen (siempre visible) → sheet row N
    //   filas[N]     = mini-cabecera de detalle        → sheet row N+1  ┐ colapsable
    //   filas[N+1..] = filas de movimiento             → sheet row N+2+ ┘
    const tz    = ss.getSpreadsheetTimeZone() || 'America/Argentina/Buenos_Aires';
    const filas = [TRAZA_GLOBAL_HEADER]; // filas[0] → sheet row 1
    const groupRanges = []; // { start: índice_mini_header_en_filas, count }

    for (const par of [...paresConMov].sort()) {
      const [codigo, lote] = par.split('|||');

      const stockIni    = getStockTotal_(mapInicio, codigo, lote);
      const stockFin    = getStockTotal_(mapFin,    codigo, lote);
      const porCaja = getDesgloseStr_(mapFinUbic, codigo, lote);

      // Fila resumen
      const resumen = new Array(TRAZA_NUM_COLS).fill('');
      resumen[TRAZA_S.CODIGO]    = codigo;
      resumen[TRAZA_S.LOTE]      = lote;
      resumen[TRAZA_S.STOCK_INI] = stockIni;
      resumen[TRAZA_S.STOCK_FIN] = stockFin;
      resumen[TRAZA_S.UBICACION] = porCaja;
      filas.push(resumen);

      const groupStart = filas.length; // índice de la mini-cabecera (primer row del grupo)

      // Mini-cabecera interna (forma parte del grupo colapsable)
      filas.push([...TRAZA_DETAIL_HEADER]);

      // Filas detalle: solo tipos relevantes, ordenadas por fecha
      const detalle = movEfectivos
        .filter(({ row, tipo }) => {
          if (!TIPOS_DETALLE_TRAZA.has(tipo)) return false;
          if (trazaNorm_(row[2]) !== codigo || trazaNorm_(row[3]) !== lote) return false;
          const f = trazaToDate_(row[0]);
          return f && f >= dInicio && f < dFinExcl;
        })
        .sort((a, b) => trazaToDate_(a.row[0]) - trazaToDate_(b.row[0]));

      for (const { row, tipo } of detalle) {
        const cantidad      = Number(row[4] || 0);
        const cajaOrigen    = (row[5] || '').toString().trim() || 'Depo';
        const idCx          = (row[10] || '').toString().trim();
        const cx            = agendaMap[idCx] || {};
        const observaciones = (row[9] || '').toString().trim();

        let signo, ubicacion;
        switch (tipo) {
          case 'ingreso':
          case 'ingreso desde liberaciones':
            signo = '+'; ubicacion = 'Depo'; break;
          case 'consumo':
            signo = '−'; ubicacion = cajaOrigen; break;
          case 'distribucion':
          case 'egreso':
            signo = '−'; ubicacion = cajaOrigen || 'Depo'; break;
          default:
            signo = ''; ubicacion = cajaOrigen;
        }

        const fechaMovFmt = (() => {
          const d = trazaToDate_(row[0]);
          return d ? Utilities.formatDate(d, tz, 'dd/MM/yyyy') : '';
        })();
        const fechaCxFmt = (() => {
          if (!cx.fecha) return '';
          const d = trazaToDate_(cx.fecha);
          return d ? Utilities.formatDate(d, tz, 'dd/MM/yyyy') : '';
        })();

        const det = new Array(TRAZA_NUM_COLS).fill('');
        det[TRAZA_D.FECHA_MOV]    = fechaMovFmt;
        det[TRAZA_D.TIPO]         = (row[1] || '').toString();
        det[TRAZA_D.CANT]         = `${signo}${cantidad}`;
        det[TRAZA_D.UBICACION]    = ubicacion;
        det[TRAZA_D.PACIENTE]     = cx.paciente    || (row[7] || '').toString().trim();
        det[TRAZA_D.CLIENTE]      = cx.cliente     || (row[8] || '').toString().trim();
        det[TRAZA_D.MEDICO]       = cx.medico      || '';
        det[TRAZA_D.INSTITUCION]  = cx.institucion || '';
        det[TRAZA_D.ID_CX]        = idCx;
        det[TRAZA_D.FECHA_CX]     = fechaCxFmt;
        det[TRAZA_D.OBSERVACIONES] = observaciones;
        // Factura: desde agenda (consumo) o extraída de observaciones (distribución)
        const facturaObs = (() => {
          const m = observaciones.match(/FV\s+[A-Z]\s+\d{4}-\d+/i);
          return m ? m[0] : '';
        })();
        det[TRAZA_D.FACTURA] = cx.factura || facturaObs;
        filas.push(det);
      }

      const groupCount = filas.length - groupStart; // mini-header + data rows
      groupRanges.push({ start: groupStart, count: groupCount });
    }

    // ── Escribir hoja ────────────────────────────────────────────────────────
    let hoja = ss.getSheetByName(nombreHoja);
    if (hoja) ss.deleteSheet(hoja);
    hoja = ss.insertSheet(nombreHoja);

    hoja.getRange(1, 1, filas.length, TRAZA_NUM_COLS).setValues(filas);

    // Cabecera global (fila 1)
    hoja.getRange(1, 1, 1, TRAZA_NUM_COLS)
      .setFontWeight('bold')
      .setBackground('#4a86c8')
      .setFontColor('#ffffff');

    // Formato por grupo:
    //   start     → sheet row = start     (fila resumen, filas[start-1])
    //   start+1   → sheet row = start+1   (mini-cabecera, filas[start])
    //   start+2.. → sheet rows            (datos de movimiento)
    for (let i = 0; i < groupRanges.length; i++) {
      const { start, count } = groupRanges[i];
      const summaryRow    = start;       // fila resumen
      const miniHeaderRow = start + 1;   // mini-cabecera
      const dataStartRow  = start + 2;   // primero dato (puede no existir si count=1)
      const dataCount     = count - 1;   // filas de datos (excluyendo mini-header)

      // Resumen: azul claro, negrita
      hoja.getRange(summaryRow, 1, 1, TRAZA_NUM_COLS)
        .setFontWeight('bold')
        .setBackground('#d9e8f7');

      // Mini-cabecera: gris, cursiva
      hoja.getRange(miniHeaderRow, 1, 1, TRAZA_NUM_COLS)
        .setFontStyle('italic')
        .setFontWeight('bold')
        .setBackground('#e8e8e8')
        .setFontColor('#555555');

      // Filas de datos: alternando blanco / azul muy claro
      if (dataCount > 0) {
        hoja.getRange(dataStartRow, 1, dataCount, TRAZA_NUM_COLS)
          .setBackground(i % 2 === 0 ? '#f4f8fd' : '#ffffff');
      }
    }

    // Agrupación: el +/− aparece ANTES de cada grupo (junto a la fila resumen)
    // El grupo incluye mini-cabecera + datos → sheet rows start+1 .. start+count
    hoja.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
    for (const { start, count } of groupRanges) {
      hoja.getRange(start + 1, 1, count, 1).shiftRowGroupDepth(1);
    }
    hoja.collapseAllRowGroups();

    hoja.autoResizeColumns(1, TRAZA_NUM_COLS);

    const totalMovs = groupRanges.reduce((s, g) => s + g.count - 1, 0); // -1 excluye mini-header
    return `✅ "${nombreHoja}" generada: ${groupRanges.length} producto/s, ${totalMovs} movimiento/s.`;

  } catch (e) {
    return '❌ Error generando trazabilidad: ' + e.message;
  }
}

// Regenera todas las hojas para todos los meses con datos
function regenerarTrazabilidadCompleta() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shMov = ss.getSheetByName('Movimientos');
  if (!shMov || shMov.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No hay movimientos registrados.');
    return;
  }

  const lastRow       = shMov.getLastRow();
  const movData       = shMov.getRange(2, 1, lastRow - 1, 11).getValues();
  const filasAnuladas = getFilasAnuladas_(movData);
  const claves        = new Set();

  for (let i = 0; i < movData.length; i++) {
    if (filasAnuladas.has(i + 2)) continue;
    const tipo = normalizarTipo_((movData[i][1] || '').toString());
    if (!TIPOS_DETALLE_TRAZA.has(tipo)) continue;
    const fecha = trazaToDate_(movData[i][0]);
    if (fecha && !isNaN(fecha)) {
      claves.add(`${fecha.getFullYear()}-${fecha.getMonth()}`);
    }
  }

  if (claves.size === 0) {
    SpreadsheetApp.getUi().alert('No hay movimientos para incluir en la Trazabilidad.');
    return;
  }

  const meses = [...claves]
    .map(k => { const [ano, mes] = k.split('-').map(Number); return { ano, mes }; })
    .sort((a, b) => a.ano !== b.ano ? a.ano - b.ano : a.mes - b.mes);

  meses.forEach(({ mes, ano }) => generarTrazabilidadDesdeMes(mes, ano));

  const etiquetas = meses.map(({ mes, ano }) => `${MESES_ES[mes]} ${ano}`).join(', ');
  SpreadsheetApp.getUi().alert(`✅ Trazabilidad regenerada para: ${etiquetas}`);
}

// ─── Helpers internos ─────────────────────────────────────────────────────────

// Normaliza código/lote a mayúsculas sin espacios
function trazaNorm_(v) {
  return (v || '').toString().trim().toUpperCase();
}

// Convierte valor de celda a Date (null si inválido)
function trazaToDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

// Lee el inventario actual desde la hoja Inventario.
// Retorna: { 'CODIGO|||LOTE': totalCantidad } sumando todas las ubicaciones.
// Usar el total evita problemas con cajaOrigen = 'N/A' o vacío en los movimientos.
function leerInventarioActual_(ss) {
  const invSheet = ss.getSheetByName('Inventario');
  if (!invSheet || invSheet.getLastRow() < 2) return {};
  // Índices 1-based según getDefaultColumnIndexInventario: codigo=2, lote=3, cantidad=4
  const COL_COD  = 2; // B
  const COL_LOTE = 3; // C
  const COL_CANT = 4; // D
  const data = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, COL_CANT).getValues();
  const map  = {};
  for (const row of data) {
    const codigo   = trazaNorm_(row[COL_COD  - 1]);
    const lote     = trazaNorm_(row[COL_LOTE - 1]);
    const cantidad = Number(row[COL_CANT - 1] || 0);
    if (!codigo || !lote) continue;
    const k = `${codigo}|||${lote}`;
    map[k] = (map[k] || 0) + cantidad;
  }
  return map;
}

// Deshace el EFECTO NETO sobre el total de stock de una lista de movimientos.
// Modifica el mapa en-place. Al operar con totales no es necesario rastrear ubicaciones,
// lo que evita falsos resultados cuando cajaOrigen es 'N/A' o está vacío.
//
// Efectos netos:
//   Ingreso / Ingreso desde lib  → +cant sobre el total  → undo = −cant
//   Consumo / Egreso / Distrib.  → −cant sobre el total  → undo = +cant
//   Reposición / Entre cajas     → 0 neto               → sin cambio
function deshacerMovimientos_(mapa, movimientos) {
  for (const { row, tipo } of movimientos) {
    const codigo   = trazaNorm_(row[2]);
    const lote     = trazaNorm_(row[3]);
    const cantidad = Number(row[4] || 0);
    if (!codigo || !lote || cantidad <= 0) continue;
    const k = `${codigo}|||${lote}`;
    switch (tipo) {
      case 'ingreso':
      case 'ingreso desde liberaciones':
        mapa[k] = (mapa[k] || 0) - cantidad; // undo ingreso: quitar lo que entró
        break;
      case 'consumo':
      case 'egreso':
      case 'distribucion':
        mapa[k] = (mapa[k] || 0) + cantidad; // undo salida: reponer lo que salió
        break;
      // reposicion / entre cajas: efecto neto = 0, no modificar total
    }
  }
}

// Obtiene el stock total para un par código/lote
function getStockTotal_(mapa, codigo, lote) {
  return mapa[`${codigo}|||${lote}`] || 0;
}

// Lee inventario actual por ubicación.
// Retorna: { 'CODIGO|||LOTE|||UBICACION': cantidad }
function leerInventarioUbicacion_(ss) {
  const invSheet = ss.getSheetByName('Inventario');
  if (!invSheet || invSheet.getLastRow() < 2) return {};
  const COL_COD  = 2; const COL_LOTE = 3; const COL_CANT = 4; const COL_UBIC = 5;
  const data = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, COL_UBIC).getValues();
  const map  = {};
  for (const row of data) {
    const codigo    = trazaNorm_(row[COL_COD  - 1]);
    const lote      = trazaNorm_(row[COL_LOTE - 1]);
    const cantidad  = Number(row[COL_CANT - 1] || 0);
    const ubicacion = trazaNorm_(row[COL_UBIC - 1]);
    if (!codigo || !lote || !ubicacion || ubicacion === 'N/A') continue;
    const k = `${codigo}|||${lote}|||${ubicacion}`;
    map[k] = (map[k] || 0) + cantidad;
  }
  return map;
}

// Deshace movimientos sobre un mapa por ubicación.
// Omite movimientos con cajaOrigen/destino='N/A' o vacío (no afectan el desglose).
function deshacerPorUbicacion_(mapa, movimientos) {
  const add = (codigo, lote, ubicacion, delta) => {
    const ub = trazaNorm_(ubicacion);
    if (!ub || ub === 'N/A') return;
    const k = `${codigo}|||${lote}|||${ub}`;
    mapa[k] = (mapa[k] || 0) + delta;
  };
  for (const { row, tipo } of movimientos) {
    const codigo   = trazaNorm_(row[2]);
    const lote     = trazaNorm_(row[3]);
    const cantidad = Number(row[4] || 0);
    if (!codigo || !lote || cantidad <= 0) continue;
    const origen  = trazaNorm_(row[5]);
    const destino = trazaNorm_(row[6]);
    switch (tipo) {
      case 'ingreso':
      case 'ingreso desde liberaciones':
        add(codigo, lote, 'DEPO', -cantidad); break;
      case 'reposicion':
      case 'reposicion caja completa':
        add(codigo, lote, origen || 'DEPO', +cantidad);
        if (destino) add(codigo, lote, destino, -cantidad); break;
      case 'entre cajas':
        add(codigo, lote, origen, +cantidad);
        if (destino) add(codigo, lote, destino, -cantidad); break;
      case 'consumo':
        add(codigo, lote, origen, +cantidad); break;
      case 'distribucion':
      case 'egreso':
        add(codigo, lote, origen || 'DEPO', +cantidad); break;
    }
  }
}

// Devuelve texto de desglose por ubicación: "DEPO: 6 | CC-02: 3 | CC-03: 3"
function getDesgloseStr_(mapaUbic, codigo, lote) {
  const prefix = `${codigo}|||${lote}|||`;
  const entries = Object.entries(mapaUbic)
    .filter(([k]) => k.startsWith(prefix))
    .map(([k, v]) => ({ ubicacion: k.slice(prefix.length), cantidad: v }))
    .filter(e => e.cantidad > 0)
    .sort((a, b) => a.ubicacion.localeCompare(b.ubicacion));
  return entries.length > 0 ? entries.map(e => `${e.ubicacion}: ${e.cantidad}`).join(' | ') : '';
}

// Construye Set con filas (1-based en sheet) que fueron anuladas
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
// Lee todas las hojas cuyo nombre empiece con "Agenda", filtra por estado "Autorizada"
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
      sheet.getRange(2, 1, lastRow - 1, 16).getValues().forEach(row => {
        const id     = (row[AGENDA_COL_ID] || '').toString().trim();
        const estado = (row[9]             || '').toString().trim();
        if (!id || estado !== 'Autorizada') return;
        map[id] = {
          fecha:       row[AGENDA_COL_FECHA],
          paciente:    (row[AGENDA_COL_PACIENTE]    || '').toString().trim(),
          institucion: (row[AGENDA_COL_INSTITUCION] || '').toString().trim(),
          medico:      (row[AGENDA_COL_MEDICO]      || '').toString().trim(),
          cliente:     (row[AGENDA_COL_CLIENTE]     || '').toString().trim(),
          factura:     (row[AGENDA_COL_FACTURA]     || '').toString().trim()
        };
      });
    });
    return map;
  } catch (e) {
    logError('Error leyendo Agenda para trazabilidad', { error: e.message });
    return {};
  }
}
