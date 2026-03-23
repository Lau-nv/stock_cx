// actualizarStockTotal implementada en Helpers.js (actualizarStockTotal)
// validarLoteManual, limpiarCeros y normalizarTipo_ están centralizados en Helpers.js

/************  ANULACIÓN DE MOVIMIENTOS  ************/
// Punto de entrada: deshacer por número de fila (1-based) en "Movimientos"
function deshacerMovimientoPorFila(nFila) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shMov = ss.getSheetByName('Movimientos');
  const invSheet = ss.getSheetByName('Inventario');

  if (!shMov) return "❌ No existe la hoja 'Movimientos'.";
  if (!invSheet) return "❌ No existe la hoja 'Inventario'.";
  if (nFila < 2 || nFila > shMov.getLastRow()) return "❌ Número de fila inválido.";

  // Leer fila a anular
  // Columnas: Fecha y Hora | Tipo | Código | Lote | Cantidad | Caja Origen | Caja Destino | Paciente | Cliente | Observaciones | ID CX
  const row = shMov.getRange(nFila, 1, 1, 11).getValues()[0];
  const tipoRaw = (row[1] || "").toString();
  const tipo = normalizarTipo_(tipoRaw);
  const codigo = (row[2] || "").toString().trim().toUpperCase();
  const lote = (row[3] || "").toString().trim().toUpperCase();
  const cantidad = Number(row[4] || 0);
  const cajaOrigen = (row[5] || "").toString().trim().toUpperCase();
  const cajaDestino = (row[6] || "").toString().trim().toUpperCase();
  const paciente = (row[7] || "").toString().trim();
  const cliente = (row[8] || "").toString().trim();
  const observacionesOriginal = (row[9] || "").toString().trim();
  const idCxOriginal = (row[10] || "N/A").toString().trim();

  if (!codigo || !lote || !(cantidad > 0)) return "❌ La fila seleccionada no contiene datos válidos para anular.";

  // Evitar doble anulación
  const yaAnulada = existeAnulacionDeFila_(shMov, nFila);
  if (yaAnulada) return `❌ La fila ${nFila} ya fue anulada previamente.`;

  // 📝 Capturar número de fila ANTES de operaciones para referencia en logs
  const filaMovimiento = shMov.getLastRow() + 1;

  // Ejecutar reversa según tipo
  let res;
  try {
    switch (tipo) {
      case 'reposicion':
      case 'reposicion caja completa':
        // Reversa: Caja -> Depo
        const datosLogRestaRep = { tipoMovimiento: 'Anulación Reposición', ubicacionOrigen: cajaDestino, ubicacionDestino: 'Depo', filaMovimiento };
        res = restarEnInventario_(invSheet, codigo, lote, cajaDestino, cantidad, datosLogRestaRep);
        if (res !== true) return res;
        const datosLogSumaRep = { tipoMovimiento: 'Anulación Reposición', ubicacionOrigen: cajaDestino, ubicacionDestino: 'Depo', filaMovimiento };
        res = sumarEnInventario_(invSheet, codigo, lote, 'Depo', cantidad, datosLogSumaRep);
        if (res !== true) return res;
        break;

      case 'entre cajas':
        // Reversa: mover de cajaDestino -> cajaOrigen
        const datosLogRestaEC = { tipoMovimiento: 'Anulación Movimiento', ubicacionOrigen: cajaDestino, ubicacionDestino: cajaOrigen, filaMovimiento };
        res = restarEnInventario_(invSheet, codigo, lote, cajaDestino, cantidad, datosLogRestaEC);
        if (res !== true) return res;
        const datosLogSumaEC = { tipoMovimiento: 'Anulación Movimiento', ubicacionOrigen: cajaDestino, ubicacionDestino: cajaOrigen, filaMovimiento };
        res = sumarEnInventario_(invSheet, codigo, lote, cajaOrigen, cantidad, datosLogSumaEC);
        if (res !== true) return res;
        break;

      case 'consumo':
        // Reversa: reponer en cajaOrigen
        const datosLogCons = { tipoMovimiento: 'Anulación Consumo', filaMovimiento };
        res = sumarEnInventario_(invSheet, codigo, lote, cajaOrigen, cantidad, datosLogCons);
        if (res !== true) return res;
        break;

      case 'distribucion':
        // Reversa: reponer en Depo
        const datosLogDist = { tipoMovimiento: 'Anulación Distribución', filaMovimiento };
        res = sumarEnInventario_(invSheet, codigo, lote, 'Depo', cantidad, datosLogDist);
        if (res !== true) return res;
        break;
        
      case 'ingreso':
      case 'ingreso desde liberaciones':
        // Reversa: restar de Depo
        const datosLogIng = { tipoMovimiento: 'Anulación Ingreso', filaMovimiento };
        res = restarEnInventario_(invSheet, codigo, lote, 'Depo', cantidad, datosLogIng);
        if (res !== true) return res;
        break;

      case 'anulacion':
      case 'anulación':
        return "❌ No se puede anular una anulación.";

      default:
        return `❌ Tipo de movimiento no soportado para deshacer: "${tipoRaw}".`;
    }
  } catch (e) {
    return "❌ Error durante la anulación: " + e.message;
  }

  // Determinar tipo de anulación específico y columnas origen/destino según tipo original
  let tipoAnulacion = "Anulación";
  let origenAnulacion = "N/A";
  let destinoAnulacion = "N/A";
  let pacienteAnulacion = "N/A";
  let clienteAnulacion = "N/A";
  
  switch (tipo) {
    case 'reposicion':
    case 'reposicion caja completa':
      tipoAnulacion = "Anulación Reposición";
      origenAnulacion = cajaDestino || "N/A"; // Sale de donde se había reposicionado
      destinoAnulacion = "DEPO"; // Vuelve a Depo
      break;
      
    case 'entre cajas':
      tipoAnulacion = "Anulación Movimiento";
      origenAnulacion = cajaDestino || "N/A"; // Sale de donde había llegado
      destinoAnulacion = cajaOrigen || "N/A"; // Vuelve de donde salió
      break;
      
    case 'consumo':
      tipoAnulacion = "Anulación Consumo";
      origenAnulacion = "N/A"; // El consumo no tiene origen físico en la reversa
      destinoAnulacion = cajaOrigen || "N/A"; // Vuelve a la caja original
      pacienteAnulacion = paciente || "N/A"; // Mantener info del paciente original
      clienteAnulacion = cliente || "N/A"; // Mantener info del cliente original
      break;
      
    case 'distribucion':
      tipoAnulacion = "Anulación Distribución";
      origenAnulacion = "N/A"; // La distribución no tiene origen físico en la reversa
      destinoAnulacion = "DEPO"; // Vuelve a Depo (o a la caja si era desde caja)
      if (cajaOrigen && cajaOrigen !== "N/A" && cajaOrigen !== "DEPO") {
        destinoAnulacion = cajaOrigen; // Si la distribución era desde una caja
      }
      clienteAnulacion = cliente || "N/A"; // Mantener info del cliente original
      break;
      
    case 'ingreso':
    case 'ingreso desde liberaciones':
      tipoAnulacion = "Anulación Ingreso";
      origenAnulacion = "DEPO"; // Sale de Depo
      destinoAnulacion = "N/A"; // No tiene destino (se resta)
      break;
  }

  // Registrar la anulación como nueva fila en "Movimientos"
  shMov.appendRow([
    new Date(), // Fecha y Hora
    tipoAnulacion, // Tipo de Movimiento (específico)
    codigo, // Código
    lote, // Lote
    cantidad, // Cantidad
    origenAnulacion, // Caja Origen (reversa del movimiento original)
    destinoAnulacion, // Caja Destino (reversa del movimiento original)
    pacienteAnulacion, // Paciente (del movimiento original si aplica)
    clienteAnulacion, // Cliente (del movimiento original si aplica)
    `ANULACIÓN de fila ${nFila} (tipo: ${tipoRaw})${observacionesOriginal ? ' - Obs: ' + observacionesOriginal : ''}`, // Observaciones
    idCxOriginal // ID CX (copiado del movimiento original)
  ]);

  // Limpieza + totales
  // limpiarCeros(['Inventario']); // ✅ Deshabilitado: ahora mantenemos lotes en 0 para trazabilidad
  if (typeof actualizarStockTotal === "function") {
    actualizarStockTotal();
  }

  return `✅ Movimiento de la fila ${nFila} anulado correctamente.`;
}/************  HELPERS COMUNES  ************/
// Busca si ya existe una fila de "Anulación" que mencione "fila N"
function existeAnulacionDeFila_(shMov, nFila) {
  const last = shMov.getLastRow();
  if (last < 2) return false;
  const data = shMov.getRange(2, 1, last - 1, 11).getValues();

  const needle = new RegExp(`\\bfila\\s+${nFila}\\b`, 'i');
  for (let i = 0; i < data.length; i++) {
    const tipo = normalizarTipo_(data[i][1]);
    const obs = (data[i][9] || "").toString();
    if ((tipo === 'anulacion' || tipo === 'anulación') && needle.test(obs)) {
      return true;
    }
  }
  return false;
}
/************  PROMPT MEJORADO PARA DESHACER  ************/
function mostrarPromptDeshacerMovimiento() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const shMov = ss.getSheetByName('Movimientos');
  if (!shMov) {
    ui.alert("❌ No existe la hoja 'Movimientos'.");
    return;
  }

  // 1) Pedir número de fila
  const r = ui.prompt(
    "Deshacer Movimiento",
    "Ingresá el número de fila en 'Movimientos' que querés anular (≥ 2):",
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Operación cancelada.");
    return;
  }
  const fila = parseInt(r.getResponseText(), 10);
  if (isNaN(fila) || fila < 2 || fila > shMov.getLastRow()) {
    ui.alert("⚠️ Ingresá un número de fila válido (mayor o igual a 2 y dentro del rango).");
    return;
  }

  // 2) Leer y validar fila
  const row = shMov.getRange(fila, 1, 1, 11).getValues()[0];
  // Columnas: Fecha y Hora | Tipo | Código | Lote | Cantidad | Caja Origen | Caja Destino | Paciente | Cliente | Observaciones | ID CX
  const fecha = row[0] instanceof Date ? row[0] : (row[0] ? new Date(row[0]) : null);
  const tipoRaw = (row[1] || "").toString();
  const tipo = normalizarTipo_(tipoRaw);
  const codigo = (row[2] || "").toString().trim().toUpperCase();
  const lote = (row[3] || "").toString().trim().toUpperCase();
  const cantidad = Number(row[4] || 0);
  const cajaOrigen = (row[5] || "").toString().trim().toUpperCase() || "N/A";
  const cajaDestino = (row[6] || "").toString().trim().toUpperCase() || "N/A";
  const paciente = (row[7] || "").toString() || "N/A";
  const cliente = (row[8] || "").toString() || "N/A";
  const obs = (row[9] || "").toString() || "N/A";

  if (!codigo || !lote || !(cantidad > 0)) {
    ui.alert("❌ La fila seleccionada no contiene datos válidos para anular.");
    return;
  }
  if (tipo === 'anulacion' || tipo === 'anulación') {
    ui.alert("❌ No se puede anular una anulación.");
    return;
  }
  if (existeAnulacionDeFila_(shMov, fila)) {
    ui.alert(`❌ La fila ${fila} ya fue anulada previamente.`);
    return;
  }

  // 3) Armar resumen + efecto de reversa
  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || "America/Argentina/Buenos_Aires";
  const fechaStr = fecha ? Utilities.formatDate(fecha, tz, "yyyy-MM-dd HH:mm") : "N/A";
  const efecto = describirReversa_(tipo, { cajaOrigen, cajaDestino });

  const resumen = "Vas a anular el siguiente movimiento:\n\n" +
    `• Fila: ${fila}\n` +
    `• Fecha: ${fechaStr}\n` +
    `• Tipo: ${tipoRaw}\n` +
    `• Código: ${codigo}\n` +
    `• Lote: ${lote}\n` +
    `• Cantidad: ${cantidad}\n` +
    `• Caja Origen: ${cajaOrigen}\n` +
    `• Caja Destino: ${cajaDestino}\n` +
    `• Paciente: ${paciente}\n` +
    `• Cliente: ${cliente}\n` +
    `• Observaciones: ${obs}\n\n` +
    "Efecto de la reversa:\n" +
    `${efecto}\n\n` +
    "¿Confirmás la anulación?";

  const confirm = ui.alert("Confirmar anulación", resumen, ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) {
    ui.alert("Operación cancelada.");
    return;
  }

  // 4) Ejecutar anulación
  const resultado = deshacerMovimientoPorFila(fila);
  ui.alert(resultado);
}
/************  DESCRIPCIÓN DE REVERSA SEGÚN TIPO  ************/
function describirReversa_(tipoNorm, ctx) {
  // ctx: { cajaOrigen, cajaDestino }
  switch (tipoNorm) {
    case 'reposicion':
    case 'reposicion caja completa':
      return `Se RESTA la cantidad en la caja DESTINO (${ctx.cajaDestino}) y se SUMA en Depo.`;
    case 'entre cajas':
      return `Se MUEVE la cantidad desde la caja DESTINO (${ctx.cajaDestino}) hacia la caja ORIGEN (${ctx.cajaOrigen}).`;
    case 'consumo':
      return `Se REPONE la cantidad en la caja ORIGEN (${ctx.cajaOrigen}).`;
    case 'distribucion':
      return `Se REPONE la cantidad en Depo.`;
    case 'ingreso':
    case 'ingreso desde liberaciones':
      return `Se RESTA la cantidad en Depo.`;
    default:
      return `Tipo no reconocido para describir reversa (${tipoNorm}).`;
  }
}

