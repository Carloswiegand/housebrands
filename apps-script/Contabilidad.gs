// =============================================================
// HOUSEBRANDS ERP — Contabilidad.gs
// Paquete mensual para la contabilidad externa (BP Consulting: Sandra y Ambrosio).
//
// contabPreparar('2026-10') hace todo para el mes contable indicado:
//   1. Busca en Gmail las facturas que llegan por correo (The Link It, BP Consulting, ABEE,
//      Shopify MemoKoala/Montana, Google Workspace, cartola BCI) y guarda los adjuntos en Drive:
//      "Contabilidad HouseBrands/AAAA-MM".
//   2. Descarga de la API de MercadoLibre las facturas y notas de crédito de ML y Mercado Pago
//      del período (PDF + XML) y el resumen de cargos.
//   3. Arma "Cierre AAAA-MM HouseBrands.xlsx": inventario del mes (ventas por canal, stock final,
//      valorización y diferencia con ABEE), ABEE (envíos + comisión), resumen y detalle de cargos ML/MP.
//   4. Deja la hoja Contab_Checklist al día (la lee HB Control → Estado del sistema).
//   5. Crea un BORRADOR en Gmail para Sandra y Ambrosio con todo adjunto. Nunca lo envía.
//
// Lo que no se puede traer solo queda en el checklist como "Manual": estado de cuenta de
// Mercado Pago, reporte de pagos de facturas ML, cartola Excel BCI y estados de cuenta de
// tarjetas (los PDF del banco vienen con clave).
//
// Hojas: Contab_Checklist (estado por mes), Contab_Inventario (cadena de inventario mes a mes),
//        Contab_ABEE (envíos y comisión por mes).
// Activador: contabProgramado el día 19 de cada mes (después de las facturas de Shopify).
// =============================================================

const CONTAB_PARA = 'sandra@bp-consulting.cl,aboul@bp-consulting.cl';
const CONTAB_CARPETA = 'Contabilidad HouseBrands';
const CONTAB_TZ = 'America/Santiago';
const CONTAB_MAX_ADJ = 23 * 1024 * 1024;   // Gmail permite 25 MB; se deja margen por la codificación

const H_CONTAB = 'Contab_Checklist';
const CONTAB_ENC = ['Mes', 'Grupo', 'Ítem', 'Estado', 'Archivos', 'Enlace', 'Nota', 'Actualizado'];
const H_CONTAB_INV = 'Contab_Inventario';
const CONTAB_INV_ENC = ['Mes', 'SKU', 'Producto', 'Inicial', 'Venta ML', 'Venta Shopify', 'Otros canales', 'Final', 'Costo', 'Valorización', 'Stock ABEE', 'Diferencia'];
const H_CONTAB_ABEE = 'Contab_ABEE';
const CONTAB_ABEE_ENC = ['Mes', 'Pedidos', 'Envíos', 'Comisión', 'Reembolso', 'Neto', 'IVA', 'Total'];

// Reglas de búsqueda en Gmail. Ventana: [desfase de meses, día] desde el mes contable.
//   adjuntar: va en el correo a Sandra. protegido: PDF con clave (se guarda, no se adjunta).
const CONTAB_CORREO = [
  { id: 'linkit', grupo: 'Proveedores', item: 'Factura The Link It (fee agencia)', q: 'from:pagos@thelinkit.com subject:factura', desde: [0, 1], hasta: [0, 28], ext: /\.pdf$/i, llega: 'primeros días del mes' },
  { id: 'bpc', grupo: 'Proveedores', item: 'Factura BP Consulting', q: 'from:aboul@bp-consulting.cl has:attachment', desde: [0, 20], hasta: [1, 12], ext: /\.pdf$/i, llega: 'fin de mes' },
  { id: 'abee', grupo: 'Proveedores', item: 'Factura y desglose ABEE', q: 'from:baltazar@abeechile.com has:attachment', desde: [1, 1], hasta: [1, 20], ext: /\.(pdf|xlsx)$/i, llega: 'primera semana del mes siguiente' },
  { id: 'shopify', grupo: 'Shopify', item: 'Facturas Shopify MemoKoala y Montana', q: 'from:billing@shopify.com subject:factura', desde: [1, 1], hasta: [1, 28], ext: /\.pdf$/i, llega: '~17 (MemoKoala) y ~18 (Montana) del mes siguiente', minimo: 2 },
  { id: 'workspace', grupo: 'Proveedores', item: 'Factura Google Workspace', q: 'from:payments-noreply@google.com subject:"Google Workspace"', desde: [1, 1], hasta: [1, 12], ext: /\.pdf$/i, llega: '~2 del mes siguiente' },
  { id: 'cartola', grupo: 'Banco', item: 'Cartola BCI (PDF del correo)', q: 'from:bcimail@bci.cl subject:cartola', desde: [1, 1], hasta: [1, 8], ext: /\.pdf$/i, protegido: true, llega: '~1 del mes siguiente' },
];

const CONTAB_MANUAL = [
  ['Banco', 'Cartola Excel BCI', 'Descargar desde Bci Pyme (www.bci.cl/pyme) con el botón Abrir Bci Pyme de HB Control.'],
  ['Banco', 'Estados de cuenta tarjeta nacional e internacional', 'Llegan por correo con clave: abrir, guardar sin clave en la carpeta del mes. Sandra saca de aquí los cargos de Meta y Google Ads.'],
  ['MercadoLibre', 'Estado de cuenta Mercado Pago', 'Descargar desde Mercado Pago → Reportes.'],
  ['MercadoLibre', 'Reporte de pagos de facturas ML', 'MercadoLibre → Facturación → Tarifas y pagos → mes → Ir al detalle → Reportes → Pagos de facturas.'],
];

// Nombre de archivo → ítem manual del checklist (para lo que se sube a la carpeta del mes).
const CONTAB_PATRONES = [
  ['Cartola Excel BCI', /cartola.*\.xlsx?$/i],
  ['Estados de cuenta tarjeta nacional e internacional', /estado.?de.?cuenta.?(nacional|internacional)|tarjeta/i],
  ['Estado de cuenta Mercado Pago', /mercado.?pago|account.?(statement|money)|release|liquidaci/i],
  ['Reporte de pagos de facturas ML', /pagos.?(de.?)?facturas/i],
  ['Reportes Excel de facturación ML (portal)', /reporte_(facturacion|notas)/i],
];

// ── Utilidades ──────────────────────────────────────────────

function _cMes_(mes) {
  if (mes && /^\d{4}-\d{2}$/.test(mes)) return mes;
  const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1);
  return Utilities.formatDate(d, CONTAB_TZ, 'yyyy-MM');
}
function _cFecha_(mes, desfase, dia) {
  const [a, m] = mes.split('-').map(Number);
  return new Date(a, m - 1 + desfase, dia);
}
function _cGmailFecha_(d) { return Utilities.formatDate(d, CONTAB_TZ, 'yyyy/MM/dd'); }
function _cNombreMes_(mes) {
  const [a, m] = mes.split('-').map(Number);
  const n = MESES_NOMBRE[m - 1];
  return n.charAt(0).toUpperCase() + n.slice(1) + ' ' + a;
}
function _cPrefijo_(mes) {   // "10 Oct26 "
  const [a, m] = mes.split('-').map(Number);
  const n = MESES_NOMBRE[m - 1];
  return mes.slice(5) + ' ' + n.charAt(0).toUpperCase() + n.slice(1, 3) + String(a).slice(2) + ' ';
}
function _cHoja_(nombre, enc) {
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.getRange('A:A').setNumberFormat('@');   // "2026-10" como texto, no como fecha
    h.appendRow(enc); h.setFrozenRows(1); h.getRange(1, 1, 1, enc.length).setFontWeight('bold');
  }
  return h;
}
function _cCarpeta_(mes) {
  const raiz = DriveApp.getFoldersByName(CONTAB_CARPETA);
  const r = raiz.hasNext() ? raiz.next() : DriveApp.createFolder(CONTAB_CARPETA);
  const sub = r.getFoldersByName(mes);
  return sub.hasNext() ? sub.next() : r.createFolder(mes);
}
// Guarda (o reemplaza) un archivo con ese nombre en la carpeta.
function _cGuardar_(carpeta, blob, nombre) {
  const ex = carpeta.getFilesByName(nombre);
  while (ex.hasNext()) ex.next().setTrashed(true);
  return carpeta.createFile(blob.setName(nombre));
}

// ── 1. Gmail ────────────────────────────────────────────────

function _cCorreo_(mes, carpeta) {
  const pre = _cPrefijo_(mes), filas = [];
  CONTAB_CORREO.forEach(r => {
    const desde = _cFecha_(mes, r.desde[0], r.desde[1]), hasta = _cFecha_(mes, r.hasta[0], r.hasta[1] + 1);
    const q = r.q + ' after:' + _cGmailFecha_(desde) + ' before:' + _cGmailFecha_(hasta);
    const guardados = [];
    GmailApp.search(q, 0, 20).forEach(t => t.getMessages().forEach(msg => {
      if (/^carlosw@/i.test(msg.getFrom().replace(/.*</, ''))) return;   // reenvíos propios
      msg.getAttachments({ includeInlineImages: false }).forEach(a => {
        if (!r.ext.test(a.getName())) return;
        const nombre = pre + a.getName();
        if (guardados.some(g => g.nombre === nombre)) return;
        const f = _cGuardar_(carpeta, a.copyBlob(), nombre);
        guardados.push({ nombre: nombre, id: f.getId(), protegido: !!r.protegido });
      });
    }));
    const llegoTodo = guardados.length >= (r.minimo || 1);
    const pendiente = !llegoTodo && new Date() < hasta;
    let estado = llegoTodo ? (r.protegido ? 'Protegido' : 'Listo') : (pendiente ? 'Pendiente' : 'Falta');
    let nota = r.protegido && llegoTodo ? 'El PDF del banco tiene clave: se guardó en Drive pero no se adjunta.' : '';
    if (!llegoTodo) nota = (guardados.length ? 'Llegó ' + guardados.length + ' de ' + (r.minimo || 1) + '. ' : '') + 'Llega ' + r.llega + '.';
    filas.push({ grupo: r.grupo, item: r.item, estado: estado, archivos: guardados, nota: nota, adjuntar: !r.protegido });
  });
  return filas;
}

// ── 2. MercadoLibre / Mercado Pago (API de facturación) ─────
// El período de facturación de ML va del ~2 al ~1 del mes siguiente; su clave es el día 1 del
// mes siguiente (clave 2026-11-01 = operaciones de octubre). Límite de la API: 5 llamadas por minuto.

function _cMlGet_(path, crudo) {
  for (let i = 0; i < 3; i++) {
    Utilities.sleep(12500);
    const r = UrlFetchApp.fetch(ML_API + path, { headers: { Authorization: 'Bearer ' + _mlToken_() }, muteHttpExceptions: true });
    const c = r.getResponseCode();
    // los id de 16+ dígitos (orden, operación) se leen como texto para no perder precisión
    if (c === 200) return crudo ? r : JSON.parse(r.getContentText().replace(/"(order_id|operation_id|shipping_id|detail_id|pack_id)":(\d{15,})/g, '"$1":"$2"'));
    if (c !== 429 && c !== 206) throw new Error('ML ' + path.split('?')[0] + ' HTTP ' + c);
    Utilities.sleep(30000);
  }
  throw new Error('ML ' + path.split('?')[0] + ' sin respuesta (límite de la API)');
}

function _cML_(mes, carpeta) {
  const clave = Utilities.formatDate(_cFecha_(mes, 1, 1), CONTAB_TZ, 'yyyy-MM-dd');
  const pref = clave.slice(0, 7) + '-';
  const filas = [], resumen = [];
  let publicidad = null;
  [['ML', 'MercadoLibre'], ['MP', 'MercadoPago']].forEach(([g, nom]) => {
    ['BILL', 'CREDIT_NOTE'].forEach(tipo => {
      const item = (tipo === 'BILL' ? 'Factura ' : 'Notas de crédito ') + nom;
      try {
        const d = _cMlGet_('/billing/integration/periods/key/' + clave + '/documents?group=' + g + '&document_type=' + tipo + '&limit=100');
        const docs = d.results || [], guardados = [];
        docs.forEach(doc => (doc.files || []).forEach(f => {
          const r = _cMlGet_('/billing/integration/legal_document/' + f.file_id, true);
          const ct = String(r.getHeaders()['Content-Type'] || '');
          const ext = /pdf/i.test(ct) ? '.pdf' : (/xml/i.test(ct) ? '.xml' : '.bin');
          const nombre = pref + tipo + '-' + nom + '-' + (f.reference_number || f.file_id) + ext;
          const arch = _cGuardar_(carpeta, r.getBlob(), nombre);
          guardados.push({ nombre: nombre, id: arch.getId(), protegido: ext !== '.pdf' });   // el XML queda en Drive, no se adjunta
        }));
        const total = docs.reduce((s, x) => s + (Number(x.amount) || 0), 0);
        filas.push({ grupo: 'MercadoLibre', item: item, estado: docs.length ? 'Listo' : (tipo === 'BILL' ? 'Falta' : 'No aplica'), archivos: guardados,
          nota: docs.length ? 'Período ' + docs[0].period.date_from + ' a ' + docs[0].period.date_to + ' · total $' + Math.round(total).toLocaleString('es-CL') : 'Sin documentos en la API para ' + clave + '.', adjuntar: true });
      } catch (e) {
        filas.push({ grupo: 'MercadoLibre', item: item, estado: 'Error', archivos: [], nota: e.message, adjuntar: false });
      }
    });
    try {
      const s = _cMlGet_('/billing/integration/periods/key/' + clave + '/summary/details?group=' + g + '&document_type=BILL');
      const b = s.bill_includes || {};
      (b.charges || []).forEach(c => resumen.push([nom, 'Cargo', c.group_description ? String(c.group_description).trim() : '', c.label, Number(c.amount) || 0]));
      if (g === 'ML') {   // publicidad del período: cargos del grupo Publicidad menos sus anulaciones (BDLIT anula CDLIT)
        const tipos = (b.charges || []).filter(c => /publicidad/i.test(String(c.group_description || ''))).map(c => String(c.type || ''));
        publicidad = (b.charges || []).filter(c => tipos.indexOf(String(c.type || '')) >= 0).reduce((t, c) => t + (Number(c.amount) || 0), 0) +
          (b.bonuses || []).filter(x => tipos.indexOf('C' + String(x.type || '').slice(1)) >= 0).reduce((t, x) => t + (Number(x.amount) || 0), 0);
      }
      (b.bonuses || []).forEach(c => resumen.push([nom, 'Bonificación', c.group_description ? String(c.group_description).trim() : '', c.label, Number(c.amount) || 0]));
      if (b.total_amount !== undefined) resumen.push([nom, 'Total facturado', '', '', Number(b.total_amount) || 0]);
    } catch (e) { resumen.push([nom, 'Error', '', e.message, 0]); }
  });
  return { filas: filas, resumen: resumen, clave: clave, publicidad: publicidad };
}

function _cMlDetalle_(clave, g) {
  const out = [];
  let offset = 0;
  for (let p = 0; p < 10; p++) {
    const d = _cMlGet_('/billing/integration/periods/key/' + clave + '/group/' + g + '/details?document_type=BILL&limit=1000&offset=' + offset);
    (d.results || []).forEach(x => {
      const ci = x.charge_info || {}, di = x.discount_info || {}, si = (x.sales_info || [])[0] || {};
      out.push([String(ci.creation_date_time || '').replace('T', ' '), ci.legal_document_number || '', ci.transaction_detail || '', ci.detail_type || '',
        Number(ci.detail_amount) || 0, Number(di.charge_amount_without_discount) || 0, Number(di.discount_amount) || 0,
        si.order_id ? String(si.order_id) : '', si.sales_channel || '', Number(si.transaction_amount) || 0]);
    });
    offset += (d.results || []).length;
    if (!d.results || !d.results.length || offset >= (d.total || 0)) break;
  }
  return out;
}

// ── 3. ABEE (desglose .xlsx) ────────────────────────────────

function _cXlsxFilas_(blob) {
  const zip = Utilities.unzip(blob.setContentType('application/zip'));
  const arch = n => zip.filter(f => f.getName() === n)[0];
  const ss = arch('xl/sharedStrings.xml'), hoja = arch('xl/worksheets/sheet1.xml');
  if (!hoja) throw new Error('desglose ABEE sin hoja');
  const strings = [];
  if (ss) {
    const raiz = XmlService.parse(ss.getDataAsString('UTF-8')).getRootElement(), ns = raiz.getNamespace();
    raiz.getChildren('si', ns).forEach(si => strings.push(si.getValue()));
  }
  const raiz = XmlService.parse(hoja.getDataAsString('UTF-8')).getRootElement(), ns = raiz.getNamespace();
  const filas = [];
  raiz.getChild('sheetData', ns).getChildren('row', ns).forEach(row => {
    const f = {};
    row.getChildren('c', ns).forEach(c => {
      const ref = c.getAttribute('r').getValue(), col = ref.replace(/\d+/g, '');
      const v = c.getChild('v', ns), fo = c.getChild('f', ns), t = c.getAttribute('t');
      let val = v ? v.getText() : '';
      if (t && t.getValue() === 's' && val !== '') val = strings[Number(val)];
      if (t && t.getValue() === 'inlineStr') { const is = c.getChild('is', ns); val = is ? is.getValue() : ''; }
      f[col] = { v: val, f: fo ? fo.getText() : '' };
    });
    filas.push(f);
  });
  return filas;
}

// Envíos = suma de "Valor Envío"; comisión según la fórmula del propio desglose
// (pedidos > $43.000: 0,05 UF c/u; resto: % sobre la venta neta).
function _cAbee_(mes, archivosAbee) {
  const x = (archivosAbee || []).filter(a => /\.xlsx$/i.test(a.nombre))[0];
  if (!x) return null;
  const filas = _cXlsxFilas_(DriveApp.getFileById(x.id).getBlob());
  let envios = 0, pedidos = 0, reembolso = 0, uf = 0, tasa = 0.055;
  const ventas = [];
  filas.forEach(f => {
    if (f.B && /^\d+$/.test(String(f.B.v)) && f.H) { pedidos++; envios += Number(f.H.v) || 0; ventas.push(Number(f.I && f.I.v) || 0); }
    Object.keys(f).forEach(k => {
      const fo = f[k].f;
      if (/COUNTIF/i.test(fo)) {
        const mu = fo.match(/\*0\.05\*([\d.]+)/); if (mu) uf = Number(mu[1]);
        const mt = fo.match(/\/1\.19\*([\d.]+)\s*$/); if (mt) tasa = Number(mt[1]);
      }
    });
    if (f.G && /reembolso/i.test(String(f.G.v)) && f.I) reembolso = Number(f.I.v) || 0;
  });
  if (!uf) throw new Error('no encontré el valor UF en el desglose ABEE');
  const comision = ventas.filter(v => v > 43000).length * 0.05 * uf + ventas.filter(v => v < 43000).reduce((s, v) => s + v, 0) / 1.19 * tasa;
  const neto = Math.round(comision - reembolso + envios);
  const fila = [mes, pedidos, Math.round(envios), Math.round(comision), Math.round(reembolso), neto, Math.round(neto * 0.19), Math.round(neto * 1.19)];
  const h = _cHoja_(H_CONTAB_ABEE, CONTAB_ABEE_ENC);
  const datos = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, 1).getValues() : [];
  const i = datos.findIndex(r => _cMesDe_(r[0]) === mes);
  if (i >= 0) { h.getRange(i + 2, 1).setNumberFormat('@'); h.getRange(i + 2, 1, 1, fila.length).setValues([fila]); } else h.appendRow(fila);
  return fila;
}

// ── 4. Inventario del mes ───────────────────────────────────
// Inicial = final del mes anterior en Contab_Inventario (si no hay, foto de Inventario_Diario).
// Ventas = unidades del Libro_Ventas del mes por canal. Stock ABEE = foto del último día del mes.

function _cFoto_(hasta) {   // { sku: {producto, abee} } de la última foto con fecha <= hasta
  const h = SpreadsheetApp.getActive().getSheetByName(H_INVENTARIO);
  if (!h || h.getLastRow() < 2) return { fecha: '', stock: {} };
  const v = h.getRange(2, 1, h.getLastRow() - 1, ENC_INVENTARIO.length).getValues().map(r => { r[0] = _cIso_(r[0]); return r; });
  const tope = Utilities.formatDate(hasta, CONTAB_TZ, 'yyyy-MM-dd');
  const fechas = v.map(r => r[0]).filter(f => f && f <= tope).sort();
  const fecha = fechas[fechas.length - 1] || '';
  const stock = {};
  v.filter(r => r[0] === fecha).forEach(r => {
    stock[String(r[1]).trim()] = { producto: r[2], abee: (Number(r[10]) || 0) - (Number(r[9]) || 0) };   // Stock total − externo
  });
  return { fecha: fecha, stock: stock };
}

// Mes 'AAAA-MM' de una celda que puede venir como texto o como fecha (filas antiguas).
function _cMesDe_(x) {
  if (x instanceof Date) return Utilities.formatDate(x, CONTAB_TZ, 'yyyy-MM');
  return String(x || '').trim().slice(0, 7);
}

// Fecha de la hoja (Date o texto "2026-11-16" / "16-11-2026" / "16/11/2026") → "2026-11-16"
function _cIso_(x) {
  if (x instanceof Date) return Utilities.formatDate(x, CONTAB_TZ, 'yyyy-MM-dd');
  const s = String(x || '').trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/); if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function _cInventario_(mes) {
  const [a, m] = mes.split('-').map(Number);
  const ant = Utilities.formatDate(_cFecha_(mes, -1, 1), CONTAB_TZ, 'yyyy-MM');
  const h = _cHoja_(H_CONTAB_INV, CONTAB_INV_ENC);
  const datos = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, CONTAB_INV_ENC.length).getValues() : [];
  const previo = {};
  datos.filter(r => _cMesDe_(r[0]) === ant).forEach(r => { previo[r[1]] = Number(r[7]) || 0; });

  const ventas = {}, nombres = {};
  const lv = SpreadsheetApp.getActive().getSheetByName('Libro_Ventas');
  const vv = lv.getRange(2, 1, Math.max(lv.getLastRow() - 1, 1), 10).getValues();
  vv.forEach(r => {
    if (Number(r[1]) !== m || Number(r[3]) !== a) return;
    const sku = String(r[7]).trim(); if (!sku) return;
    if (r[8]) nombres[sku] = String(r[8]);
    const o = ventas[sku] = ventas[sku] || { ml: 0, sh: 0, otro: 0 };
    const q = Number(r[9]) || 0, canal = String(r[5]);
    if (/mercado|^ml$/i.test(canal)) o.ml += q; else if (/shopify/i.test(canal)) o.sh += q; else o.otro += q;
  });

  const fotoIni = _cFoto_(_cFecha_(mes, 0, 0)), fotoFin = _cFoto_(_cFecha_(mes, 1, 0));
  const skus = Object.keys(COSTOS_SKU).filter(s => previo[s] !== undefined || ventas[s] || fotoFin.stock[s] || fotoIni.stock[s]);
  const filas = skus.map(sku => {
    const v = ventas[sku] || { ml: 0, sh: 0, otro: 0 };
    const ini = previo[sku] !== undefined ? previo[sku] : ((fotoIni.stock[sku] || {}).abee || 0);
    const fin = ini - v.ml - v.sh - v.otro, costo = Number(COSTOS_SKU[sku]) || 0;
    const abee = fotoFin.stock[sku] ? fotoFin.stock[sku].abee : '';
    const prod = (fotoFin.stock[sku] || fotoIni.stock[sku] || {}).producto || nombres[sku] || '';
    return [mes, sku, prod, ini, v.ml, v.sh, v.otro, fin, costo, fin * costo, abee, abee === '' ? '' : abee - fin];
  });
  // reemplaza las filas del mes
  for (let i = datos.length - 1; i >= 0; i--) if (_cMesDe_(datos[i][0]) === mes) h.deleteRow(i + 2);
  if (filas.length) h.getRange(h.getLastRow() + 1, 1, filas.length, CONTAB_INV_ENC.length).setValues(filas);
  return { filas: filas, fotoFin: fotoFin.fecha, inicialDesde: Object.keys(previo).length ? 'cierre de ' + ant : 'foto de inventario ' + (fotoIni.fecha || '–') };
}

// ── 5. Excel del cierre ─────────────────────────────────────

function _cExcel_(mes, carpeta, inv, abeeFila, ml) {
  const tmp = SpreadsheetApp.create('tmp cierre ' + mes);
  try {
    const fmt = (h, n, cols) => { if (n > 0) h.getRange(2, cols[0], n, cols[1]).setNumberFormat('#,##0'); };
    const nombreMes = _cNombreMes_(mes);

    const hi = tmp.getSheets()[0].setName('Inventario ' + nombreMes);
    const encI = ['SKU', 'Producto', 'Inventario inicial', 'Venta ML', 'Venta Shopify', 'Otros canales', 'Inventario final', 'Costo unitario', 'Valorización stock', 'Stock ABEE (' + (inv.fotoFin || '–') + ')', 'Diferencia ABEE − libro'];
    const filasI = inv.filas.map((r, i) => {
      const n = i + 2;
      return [r[1], r[2], r[3], r[4], r[5], r[6], '=C' + n + '-D' + n + '-E' + n + '-F' + n, r[8], '=G' + n + '*H' + n, r[10], r[10] === '' ? '' : '=J' + n + '-G' + n];
    });
    hi.getRange(1, 1, 1, encI.length).setValues([encI]).setFontWeight('bold');
    if (filasI.length) {
      hi.getRange(2, 1, filasI.length, encI.length).setValues(filasI);
      const t = filasI.length + 2;
      hi.getRange(t, 1, 1, encI.length).setValues([['Total', '', '=SUM(C2:C' + (t - 1) + ')', '=SUM(D2:D' + (t - 1) + ')', '=SUM(E2:E' + (t - 1) + ')', '=SUM(F2:F' + (t - 1) + ')', '=SUM(G2:G' + (t - 1) + ')', '', '=SUM(I2:I' + (t - 1) + ')', '=SUM(J2:J' + (t - 1) + ')', '=SUM(K2:K' + (t - 1) + ')']]).setFontWeight('bold');
      fmt(hi, filasI.length + 1, [3, 9]);
    }
    hi.getRange(filasI.length + 4, 1).setValue('Inventario inicial desde: ' + inv.inicialDesde + '. Ventas en unidades desde el Libro de ventas. Stock ABEE: foto del stock publicado al ' + (inv.fotoFin || '–') + '.');

    const ha = tmp.insertSheet('ABEE');
    const abee = _cHoja_(H_CONTAB_ABEE, CONTAB_ABEE_ENC);
    const abeeMes = abee.getLastRow() > 1 ? abee.getRange(2, 1, abee.getLastRow() - 1, 1).getValues() : [];
    const hist = abee.getLastRow() > 1 ? abee.getRange(2, 1, abee.getLastRow() - 1, CONTAB_ABEE_ENC.length).getDisplayValues().map((r, k, a) => [_cMesDe_(abeeMes[k][0])].concat(r.slice(1))).filter(r => r[0] <= mes).sort((p, q) => p[0] < q[0] ? -1 : 1) : [];
    ha.getRange('A:A').setNumberFormat('@');
    ha.getRange(1, 1, 1, CONTAB_ABEE_ENC.length).setValues([CONTAB_ABEE_ENC]).setFontWeight('bold');
    if (hist.length) {
      const num = s => Number(String(s).replace(/\./g, '').replace(',', '.')) || 0;
      ha.getRange(2, 1, hist.length, CONTAB_ABEE_ENC.length).setValues(hist.map(r => [r[0]].concat(r.slice(1).map(num))));
      fmt(ha, hist.length, [2, 7]);
    }
    if (!abeeFila) ha.getRange(hist.length + 3, 1).setValue('Falta el desglose de ABEE de ' + nombreMes + '.');

    const hr = tmp.insertSheet('Resumen ML y MP');
    hr.getRange(1, 1, 1, 5).setValues([['Cuenta', 'Tipo', 'Grupo', 'Concepto', 'Monto']]).setFontWeight('bold');
    if (ml.resumen.length) { hr.getRange(2, 1, ml.resumen.length, 5).setValues(ml.resumen); hr.getRange(2, 5, ml.resumen.length, 1).setNumberFormat('#,##0'); }
    hr.getRange(ml.resumen.length + 3, 1).setValue('Período de facturación MercadoLibre ' + ml.clave + ' (operaciones de ' + nombreMes + '). Fuente: API de facturación de MercadoLibre.');

    [['ML', 'Detalle cargos ML'], ['MP', 'Detalle cargos MP']].forEach(([g, nom]) => {
      const hd = tmp.insertSheet(nom);
      const enc = ['Fecha', 'Documento', 'Detalle', 'Tipo', 'Monto cargo', 'Cargo sin descuento', 'Descuento', 'Orden', 'Canal', 'Monto venta'];
      hd.getRange('B:B').setNumberFormat('@'); hd.getRange('H:H').setNumberFormat('@');   // n° de orden sin redondear
      hd.getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight('bold');
      let det = [];
      try { det = _cMlDetalle_(ml.clave, g); } catch (e) { hd.getRange(3, 1).setValue('No se pudo leer el detalle: ' + e.message); }
      if (det.length) { hd.getRange(2, 1, det.length, enc.length).setValues(det); hd.getRange(2, 5, det.length, 3).setNumberFormat('#,##0'); hd.getRange(2, 10, det.length, 1).setNumberFormat('#,##0'); }
    });
    SpreadsheetApp.flush();

    const url = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx';
    const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
    const nombre = 'Cierre ' + mes + ' HouseBrands.xlsx';
    const f = _cGuardar_(carpeta, blob, nombre);
    return { nombre: nombre, id: f.getId(), protegido: false };
  } finally {
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
  }
}

// ── 5b. Gastos del mes al P&L ───────────────────────────────
// Solo lo que sale exacto del sistema: ABEE (desglose) y publicidad de MercadoLibre (API de facturación;
// las facturas de ML incluyen IVA, se carga el neto). No pisa lo que ya esté cargado a mano: si el concepto
// ya tiene valor ese mes en Gastos_Variables, no se toca. Meta Ads, Google Ads y The Link It quedan fuera.

function _cGastosPnl_(mes, abeeFila, ml) {
  const [a, m] = mes.split('-').map(Number);
  const nomMes = MESES_NOMBRE[m - 1];
  const fecha = Utilities.formatDate(_cFecha_(mes, 1, 0), CONTAB_TZ, 'yyyy-MM-dd');
  const hv = SpreadsheetApp.getActive().getSheetByName(API_HOJAS.gastos);
  const enc = hv ? hv.getRange(1, 1, 1, hv.getLastColumn()).getValues()[0].map(x => String(x).trim()) : [];
  const fila = hv && hv.getLastRow() > 1 ? hv.getRange(2, 1, hv.getLastRow() - 1, hv.getLastColumn()).getValues()
    .find(r => String(r[1]).trim().toLowerCase() === nomMes && Number(r[2]) === a) : null;
  const yaTiene = c => { const i = enc.indexOf(c); return !!(fila && i >= 0 && fila[i] !== '' && fila[i] !== null && Number(fila[i]) !== 0); };
  const cargas = [];
  if (abeeFila) {
    cargas.push({ concepto: 'ABEE Fulfillment', neto: abeeFila[3] - abeeFila[4], proveedor: 'ABEE SpA', documento: 'ABEE-' + mes + '-comision', nota: 'Comisión calculada del desglose ABEE (' + abeeFila[1] + ' pedidos).' });
    cargas.push({ concepto: 'ABEE Envíos', neto: abeeFila[2], proveedor: 'ABEE SpA', documento: 'ABEE-' + mes + '-envios', nota: 'Suma de envíos del desglose ABEE.' });
  }
  if (ml && ml.publicidad > 0) {
    const neto = Math.round(ml.publicidad / 1.19);
    cargas.push({ concepto: 'ML Ads', neto: neto, iva: Math.round(ml.publicidad) - neto, proveedor: 'MercadoLibre', documento: 'ML-ADS-' + ml.clave, nota: 'Publicidad del período de facturación ' + ml.clave + ' (total con IVA $' + Math.round(ml.publicidad).toLocaleString('es-CL') + ').' });
  }
  const hechos = [], omitidos = [];
  cargas.forEach(c => {
    if (!(c.neto > 0)) return;
    if (yaTiene(c.concepto)) { omitidos.push(c.concepto + ' (ya cargado)'); return; }
    try {
      gastoDocAgregar_({ concepto: c.concepto, fecha: fecha, mes: nomMes, anio: a, neto: c.neto, iva: c.iva !== undefined ? c.iva : Math.round(c.neto * 0.19),
        proveedor: c.proveedor, documento: c.documento, origen: 'Contabilidad', nota: c.nota });
      hechos.push(c.concepto + ' $' + Math.round(c.neto).toLocaleString('es-CL'));
    } catch (e) { omitidos.push(c.concepto + ' (' + e.message + ')'); }
  });
  return { hechos: hechos, omitidos: omitidos };
}

// ── 6. Checklist y borrador ─────────────────────────────────

function _cChecklist_(mes, filas, carpeta) {
  const h = _cHoja_(H_CONTAB, CONTAB_ENC);
  const datos = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, 1).getDisplayValues() : [];
  for (let i = datos.length - 1; i >= 0; i--) if (datos[i][0] === mes) h.deleteRow(i + 2);
  const ahora = Utilities.formatDate(new Date(), CONTAB_TZ, 'yyyy-MM-dd HH:mm');
  const salida = filas.map(f => [mes, f.grupo, f.item, f.estado, (f.archivos || []).map(a => a.nombre).join(' | '),
    f.enlace || ((f.archivos || []).length === 1 ? 'https://drive.google.com/file/d/' + f.archivos[0].id + '/view' : carpeta.getUrl()), f.nota || '', ahora]);
  if (salida.length) h.getRange(h.getLastRow() + 1, 1, salida.length, CONTAB_ENC.length).setValues(salida);
}

function _cBorrador_(mes, filas, notaExtra) {
  const adj = [];
  let peso = 0;
  const omitidos = [];
  filas.filter(f => f.adjuntar).forEach(f => (f.archivos || []).forEach(a => {
    if (a.protegido) return;
    const b = DriveApp.getFileById(a.id).getBlob();
    if (peso + b.getBytes().length > CONTAB_MAX_ADJ) { omitidos.push(a.nombre); return; }
    peso += b.getBytes().length;
    adj.push(b.setName(a.nombre));
  }));
  const faltan = filas.filter(f => ['Falta', 'Pendiente', 'Manual', 'Error'].indexOf(f.estado) >= 0)
    .map(f => '- ' + f.item + (f.estado === 'Pendiente' ? ' (llega más adelante, la reenvío)' : ''));
  const nombre = _cNombreMes_(mes);
  const cuerpo = 'Hola Sandra y Ambrosio,\n\n' +
    'Les envío la información de ' + nombre.toLowerCase() + ' de Housebrands Commerce SpA. ' +
    'El Excel "Cierre ' + mes + '" trae el inventario del mes, ABEE y el detalle de cargos de MercadoLibre y Mercado Pago.\n' +
    (notaExtra ? '\n' + notaExtra + '\n' : '') +
    (faltan.length ? '\nQueda pendiente:\n' + faltan.join('\n') + '\n' : '') +
    (omitidos.length ? '\nNo alcanzaron a ir adjuntos (límite de tamaño): ' + omitidos.join(', ') + '\n' : '') +
    '\nSaludos,\nCarlos';
  const asunto = 'Información ' + nombre + ' — Housebrands Commerce SpA';
  GmailApp.getDrafts().filter(x => x.getMessage().getSubject() === asunto).forEach(x => x.deleteDraft());   // al volver a correr, reemplaza el borrador
  const d = GmailApp.createDraft(CONTAB_PARA, asunto, cuerpo, { attachments: adj });
  return { id: d.getId(), adjuntos: adj.length, mb: Math.round(peso / 1048576 * 10) / 10, omitidos: omitidos };
}

// ── Entradas ────────────────────────────────────────────────

// notaExtra: texto opcional que va en el correo (p. ej. un pago duplicado).
function contabPreparar(mes, notaExtra) {
  mes = _cMes_(mes);
  const props = PropertiesService.getScriptProperties();
  if (notaExtra) props.setProperty('CONTAB_NOTA_' + mes, notaExtra); else notaExtra = props.getProperty('CONTAB_NOTA_' + mes) || '';
  const carpeta = _cCarpeta_(mes);
  const filas = _cCorreo_(mes, carpeta);

  let ml = { filas: [], resumen: [], clave: '' };
  try { ml = _cML_(mes, carpeta); } catch (e) { ml.filas.push({ grupo: 'MercadoLibre', item: 'Facturación MercadoLibre y Mercado Pago', estado: 'Error', archivos: [], nota: e.message }); }
  filas.push.apply(filas, ml.filas);

  let abeeFila = null;
  try { abeeFila = _cAbee_(mes, (filas.filter(f => f.item.indexOf('ABEE') >= 0)[0] || {}).archivos); }
  catch (e) { filas.push({ grupo: 'Excel', item: 'Cálculo ABEE', estado: 'Error', archivos: [], nota: e.message }); }

  try {
    const inv = _cInventario_(mes);
    const x = _cExcel_(mes, carpeta, inv, abeeFila, ml);
    filas.push({ grupo: 'Excel', item: 'Excel cierre (inventario, ABEE, cargos ML/MP)', estado: 'Listo', archivos: [x], adjuntar: true,
      nota: inv.filas.length + ' SKU · stock ABEE al ' + (inv.fotoFin || '–') + (abeeFila ? ' · ABEE neto $' + abeeFila[5].toLocaleString('es-CL') : ' · falta desglose ABEE') });
  } catch (e) { filas.push({ grupo: 'Excel', item: 'Excel cierre', estado: 'Error', archivos: [], nota: e.message }); }

  try {
    const gp = _cGastosPnl_(mes, abeeFila, ml);
    filas.push({ grupo: 'P&L', item: 'Gastos cargados al P&L', estado: 'Listo', archivos: [], adjuntar: false, enlace: '',
      nota: (gp.hechos.length ? 'Cargado: ' + gp.hechos.join(', ') : 'Nada nuevo') + (gp.omitidos.length ? ' · Sin tocar: ' + gp.omitidos.join(', ') : '') + '. Meta Ads, Google Ads y The Link It: los carga la tarea mensual o a mano.' });
  } catch (e) { filas.push({ grupo: 'P&L', item: 'Gastos cargados al P&L', estado: 'Error', archivos: [], nota: e.message }); }

  CONTAB_MANUAL.forEach(([g, it, n]) => filas.push({ grupo: g, item: it, estado: 'Manual', archivos: [], nota: n }));

  const b = _cBorrador_(mes, filas, notaExtra);
  filas.push({ grupo: 'Correo', item: 'Borrador para Sandra y Ambrosio', estado: 'Listo', archivos: [], adjuntar: false,
    enlace: 'https://mail.google.com/mail/u/0/#drafts', nota: b.adjuntos + ' adjuntos · ' + b.mb + ' MB' + (b.omitidos.length ? ' · no cupieron: ' + b.omitidos.join(', ') : '') + '. Revisar y enviar.' });
  _cChecklist_(mes, filas, carpeta);
  return { mes: mes, carpeta: carpeta.getUrl(), borrador: b };
}

// ── Documentos bajados desde el navegador (Bci, Mercado Pago) ──
// HB Control abre el portal; el archivo llega por postMessage y se guarda aquí en la carpeta del mes.
// d: { mes:'2026-10', item:<ítem manual del checklist>, nombre, tipo (MIME), b64 }

function contabSubir_(d) {
  const mes = String(d.mes || '');
  if (!/^\d{4}-\d{2}$/.test(mes)) throw new Error('mes inválido');
  const item = String(d.item || '');
  if (!CONTAB_MANUAL.some(m => m[1] === item)) throw new Error('ítem desconocido: ' + item);
  const nombre = String(d.nombre || '').replace(/[\\/:*?"<>|]/g, '-').trim().slice(0, 120);
  if (!nombre || !d.b64) throw new Error('falta el archivo');
  const bytes = Utilities.base64Decode(String(d.b64));
  if (bytes.length > 20 * 1024 * 1024) throw new Error('archivo de más de 20 MB');
  const final = _cPrefijo_(mes) + nombre;
  const f = _cGuardar_(_cCarpeta_(mes), Utilities.newBlob(bytes, String(d.tipo || 'application/octet-stream'), final), final);

  const h = _cHoja_(H_CONTAB, CONTAB_ENC);
  const v = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, CONTAB_ENC.length).getDisplayValues() : [];
  const i = v.findIndex(r => r[0] === mes && r[2] === item);
  const ahora = Utilities.formatDate(new Date(), CONTAB_TZ, 'yyyy-MM-dd HH:mm');
  if (i >= 0) {
    const previos = String(v[i][4] || '').split(' | ').filter(n => n && n !== final);
    const lista = previos.concat([final]);
    h.getRange(i + 2, 4, 1, 5).setValues([['Listo', lista.join(' | '), lista.length === 1 ? 'https://drive.google.com/file/d/' + f.getId() + '/view' : _cCarpeta_(mes).getUrl(), 'Bajado desde el portal.', ahora]]);
  }
  return { ok: true, nombre: final, bytes: bytes.length };
}

// Re-arma el borrador del mes con lo que hay en el checklist (después de subir documentos a mano).
function contabBorrador_(mes) {
  if (!/^\d{4}-\d{2}$/.test(String(mes || ''))) throw new Error('mes inválido');
  const carpeta = _cCarpeta_(mes);
  const ids = {};
  const it = carpeta.getFiles();
  while (it.hasNext()) { const f = it.next(); ids[f.getName()] = f.getId(); }
  const h = _cHoja_(H_CONTAB, CONTAB_ENC);
  const v = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, CONTAB_ENC.length).getDisplayValues() : [];
  const filas = v.filter(r => r[0] === mes && r[1] !== 'Correo').map(r => ({
    grupo: r[1], item: r[2], estado: r[3], nota: r[6], enlace: r[5], adjuntar: true,
    archivos: String(r[4] || '').split(' | ').filter(n => n && ids[n]).map(n => ({ nombre: n, id: ids[n], protegido: /\.xml$/i.test(n) || r[3] === 'Protegido' })),
  }));
  if (!filas.length) throw new Error('el mes ' + mes + ' no se ha preparado');
  // Archivos que se agregaron a la carpeta a mano (o bajados del banco / Mercado Pago): se asignan a su ítem por nombre.
  const usados = new Set([].concat.apply([], filas.map(f => f.archivos.map(x => x.nombre))));
  Object.keys(ids).filter(n => !usados.has(n) && !/\.xml$/i.test(n) && !/^Cierre /.test(n)).forEach(n => {
    const regla = CONTAB_PATRONES.filter(p => p[1].test(n))[0];
    const item = regla ? regla[0] : 'Otros documentos';
    let f = filas.filter(x => x.item === item)[0];
    if (!f) { f = { grupo: 'Otros', item: item, estado: 'Listo', nota: '', enlace: carpeta.getUrl(), adjuntar: true, archivos: [] }; filas.push(f); }
    f.archivos.push({ nombre: n, id: ids[n], protegido: false });
    f.estado = 'Listo'; f.nota = 'Agregado a la carpeta del mes.';
    f.enlace = f.archivos.length === 1 ? 'https://drive.google.com/file/d/' + ids[n] + '/view' : carpeta.getUrl();
  });
  const b = _cBorrador_(mes, filas, PropertiesService.getScriptProperties().getProperty('CONTAB_NOTA_' + mes) || '');
  filas.push({ grupo: 'Correo', item: 'Borrador para Sandra y Ambrosio', estado: 'Listo', archivos: [], adjuntar: false,
    enlace: 'https://mail.google.com/mail/u/0/#drafts', nota: b.adjuntos + ' adjuntos · ' + b.mb + ' MB' + (b.omitidos.length ? ' · no cupieron: ' + b.omitidos.join(', ') : '') + '. Revisar y enviar.' });
  _cChecklist_(mes, filas, carpeta);
  return { ok: true, adjuntos: b.adjuntos, mb: b.mb };
}

// Carga al P&L los gastos de un mes ya preparado (ABEE desde Contab_ABEE; ML Ads re-consultando el resumen de la API).
function contabGastosMes(mes) {
  mes = _cMes_(mes);
  const h = _cHoja_(H_CONTAB_ABEE, CONTAB_ABEE_ENC);
  const abee = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, CONTAB_ABEE_ENC.length).getValues().find(r => _cMesDe_(r[0]) === mes) : null;
  const clave = Utilities.formatDate(_cFecha_(mes, 1, 1), CONTAB_TZ, 'yyyy-MM-dd');
  const s = _cMlGet_('/billing/integration/periods/key/' + clave + '/summary/details?group=ML&document_type=BILL');
  const b = s.bill_includes || {};
  const tipos = (b.charges || []).filter(c => /publicidad/i.test(String(c.group_description || ''))).map(c => String(c.type || ''));
  const publicidad = (b.charges || []).filter(c => tipos.indexOf(String(c.type || '')) >= 0).reduce((t, c) => t + (Number(c.amount) || 0), 0) +
    (b.bonuses || []).filter(x => tipos.indexOf('C' + String(x.type || '').slice(1)) >= 0).reduce((t, x) => t + (Number(x.amount) || 0), 0);
  const r = _cGastosPnl_(mes, abee ? abee.map((v, i) => i === 0 ? v : Number(v) || 0) : null, { publicidad: publicidad, clave: clave });
  console.log(mes + ': ' + JSON.stringify(r));
  return r;
}

function contabProgramado() { correr_('Contabilidad', () => contabPreparar()); }

function activarContabilidad() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'contabProgramado').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('contabProgramado').timeBased().onMonthDay(19).atHour(8).inTimezone(CONTAB_TZ).create();
}
