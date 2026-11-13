// =============================================================
// HOUSEBRANDS ERP — Gastos.gs
// Documentos de gasto (facturas, boletas, comprobantes) que alimentan el P&L.
// Hoja Gastos_Detalle: una fila por documento.
//   ID | Fecha | Mes | Año | Concepto | Proveedor | RUT | Documento | Neto | IVA | Total | Origen | Nota | Cargado
// Al agregar o borrar un documento se recalcula el total del concepto en Gastos_Variables
// (suma de los netos del mes) y se regenera el P&L.
// Hoja Metas: proyección del flujo de caja por mes y escenario (la carga HB Control).
// =============================================================

const H_GDET = 'Gastos_Detalle';
const GDET_ENC = ['ID', 'Fecha', 'Mes', 'Año', 'Concepto', 'Proveedor', 'RUT', 'Documento', 'Neto', 'IVA', 'Total', 'Origen', 'Nota', 'Cargado'];
const H_METAS = 'Metas';
const METAS_ENC = ['Mes', 'Escenario', 'Ventas', 'Costo %', 'Logística %', 'Publicidad', 'Agencia', 'Contabilidad', 'Shopify', '3PL', 'Sueldo', 'Cuota'];

function _hojaGdet_() {
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(H_GDET);
  if (!h) { h = ss.insertSheet(H_GDET); h.appendRow(GDET_ENC); h.setFrozenRows(1); h.getRange(1, 1, 1, GDET_ENC.length).setFontWeight('bold'); }
  return h;
}

function gastoDocAgregar_(d) {
  const concepto = String(d.concepto || '').trim();
  if (GASTOS_VARIABLES.indexOf(concepto) === -1) throw new Error('concepto no válido: ' + concepto);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d.fecha || ''))) throw new Error('fecha inválida');
  const neto = Math.round(Number(d.neto) || 0);
  if (!(neto > 0)) throw new Error('el monto neto debe ser mayor a 0');
  const iva = Math.round(Number(d.iva) || 0), total = Math.round(Number(d.total) || (neto + iva));
  // El mes del gasto puede ser distinto al de la fecha del documento (p. ej. factura de octubre emitida en noviembre).
  const [fa, fm] = d.fecha.split('-').map(Number);
  const mes = String(d.mes || MESES_NOMBRE[fm - 1]).toLowerCase(), anio = Number(d.anio || fa);
  if (MESES_NOMBRE.indexOf(mes) === -1) throw new Error('mes inválido');
  const h = _hojaGdet_();
  const datos = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, GDET_ENC.length).getValues() : [];
  const docu = String(d.documento || '').trim(), rut = String(d.rut || '').trim();
  if (docu && datos.some(r => String(r[7]) === docu && String(r[6]) === rut && r[4] === concepto)) throw new Error('ese documento ya está cargado para ' + concepto);
  const id = 'G' + Utilities.formatDate(new Date(), 'America/Santiago', 'yyyyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  h.appendRow([id, d.fecha, mes, anio, concepto, String(d.proveedor || '').slice(0, 80), rut, docu, neto, iva, total, String(d.origen || 'Manual').slice(0, 20), String(d.nota || '').slice(0, 200), new Date()]);
  _gastoRecalcular_(mes, anio, concepto);
  return { ok: true, id: id };
}

function gastoDocBorrar_(id) {
  const h = _hojaGdet_();
  const datos = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, GDET_ENC.length).getValues() : [];
  const i = datos.findIndex(r => String(r[0]) === String(id));
  if (i === -1) throw new Error('no se encontró el documento');
  const r = datos[i];
  h.deleteRow(i + 2);
  _gastoRecalcular_(r[2], r[3], r[4]);
  return { ok: true };
}

// Total del concepto en el mes = suma de los netos de sus documentos.
function _gastoRecalcular_(mes, anio, concepto) {
  const h = _hojaGdet_();
  const datos = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, GDET_ENC.length).getValues() : [];
  const docs = datos.filter(r => String(r[2]).toLowerCase() === String(mes).toLowerCase() && Number(r[3]) === Number(anio) && r[4] === concepto);
  const total = docs.reduce((s, r) => s + (Number(r[8]) || 0), 0);
  const valores = {}; valores[concepto] = docs.length ? total : '';
  gastosGuardar_(mes, anio, valores);
}

// Reemplaza la hoja Metas completa. filas = [[mes 'yyyy-mm', escenario, ventas, costo%, log%, publicidad, agencia, contab, shopify, 3pl, sueldo, cuota], ...]
function metasGuardar_(filas) {
  if (!Array.isArray(filas) || !filas.length) throw new Error('sin filas');
  filas.forEach(f => {
    if (!/^\d{4}-\d{2}$/.test(String(f[0])) || ['Optimista', 'Neutro', 'Pesimista'].indexOf(f[1]) === -1 || f.length !== METAS_ENC.length) throw new Error('fila inválida: ' + JSON.stringify(f).slice(0, 80));
  });
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(H_METAS);
  if (!h) h = ss.insertSheet(H_METAS); else h.clearContents();
  h.getRange(1, 1, 1, METAS_ENC.length).setValues([METAS_ENC]).setFontWeight('bold');
  h.getRange(2, 1, filas.length, METAS_ENC.length).setNumberFormat('@').setValues(filas.map(f => f.map(String)));
  h.setFrozenRows(1);
  return { ok: true, filas: filas.length };
}
