// =============================================================
// HOUSEBRANDS ERP — Api.gs
// API privada que consume el Worker de Cloudflare (housebrands.cl/erp/api/*).
// Reemplaza la publicación web de las hojas (CSV públicos).
// La clave se guarda en Propiedades del script (API_KEY), nunca en el código.
// Para rotarla: ejecutar generarApiKey() y actualizar el secreto SHEETS_API_KEY en Cloudflare.
//
// GET  ?sheet=ventas|pnl|conv|inv&key=...   → CSV de la hoja
// POST {key, action:'venta_manual', venta}  → agrega una venta manual al Libro_Ventas
// POST {key, action:'anular_manual', orden, canal} → mueve una venta manual a Ventas_Anuladas
// =============================================================

const API_HOJAS = { ventas: 'Libro_Ventas', pnl: 'P&L', conv: 'Conversion', inv: 'Inventario_Diario' };
const CANALES_MANUALES = ['Venta Directa', 'HigoKids', 'Venta Corporativa', 'B2B', 'Otro'];

function doGet(e) {
  const clave = PropertiesService.getScriptProperties().getProperty('API_KEY');
  const p = (e && e.parameter) || {};
  if (p.code && p.state) return mlCallback_(p);   // retorno de la autorización de MercadoLibre (ML.gs)
  if (!clave || p.key !== clave) return texto_('forbidden');
  const nombre = API_HOJAS[p.sheet];
  if (!nombre) return texto_('unknown sheet');
  const hoja = SpreadsheetApp.getActive().getSheetByName(nombre);
  if (!hoja) return texto_('missing sheet');
  const filas = hoja.getDataRange().getDisplayValues();
  const csv = filas.map(f => f.map(celdaCsv_).join(',')).join('\n');
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
}

function doPost(e) {
  let body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'json inválido' }); }
  const clave = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!clave || body.key !== clave) return json_({ ok: false, error: 'forbidden' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return json_({ ok: false, error: 'ocupado, intenta de nuevo' });
  try {
    if (body.action === 'venta_manual') return json_(ventaManual_(body.venta || {}));
    if (body.action === 'anular_manual') return json_(anularManual_(body.orden, body.canal));
    return json_({ ok: false, error: 'acción desconocida' });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// Venta manual (boleta, factura B2B, marketplace sin API). Precios con IVA.
// venta = { fecha:'yyyy-mm-dd', canal, documento, cliente, comisionPct, envio, lineas:[{sku, qty, precioUnit}] }
function ventaManual_(v) {
  const canal = String(v.canal || '').trim();
  if (CANALES_MANUALES.indexOf(canal) === -1) throw new Error('canal no válido para venta manual: ' + canal);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(v.fecha || ''))) throw new Error('fecha inválida');
  const lineas = (v.lineas || []).filter(l => l && l.sku);
  if (!lineas.length) throw new Error('la venta no tiene productos');
  lineas.forEach(l => {
    if (!COSTOS_SKU[l.sku]) throw new Error('SKU sin costo en el ERP: ' + l.sku + ' (cargarlo en COSTOS_SKU antes de vender)');
    if (!(Number(l.qty) > 0) || !(Number(l.precioUnit) > 0)) throw new Error('cantidad o precio inválido en ' + l.sku);
  });

  const doc = String(v.documento || '').trim();
  const orden = doc ? doc : 'M-' + Utilities.formatDate(new Date(), 'America/Santiago', 'yyyyMMddHHmmss');
  const hojaV = SpreadsheetApp.getActive().getSheetByName(H_VENTAS);
  const claves = _claves(hojaV);
  const dup = lineas.find(l => claves.has(orden + '|' + canal + '|' + l.sku));
  if (dup) throw new Error('ya existe ' + orden + ' en ' + canal + ' con ' + dup.sku);

  const [a, m] = v.fecha.split('-').map(Number);
  const pct = Math.max(0, Number(v.comisionPct) || 0) / 100;
  const envio = Math.round(Number(v.envio) || 0);
  const filas = lineas.map((l, i) => {
    const qty = Number(l.qty);
    const pu = Math.round(Number(l.precioUnit));
    const bruto = pu * qty;
    const comision = Math.round(bruto * pct);
    const costoUnit = COSTOS_SKU[l.sku];
    const costoTotal = costoUnit * qty;
    const margen = bruto - costoTotal * IVA;
    return [
      v.fecha, m, MESES[m - 1], a,
      _marcaML(l.sku, NOMBRES_SKU[l.sku] || ''), canal, orden, l.sku,
      NOMBRES_SKU[l.sku] || l.sku, qty,
      pu, bruto,
      i === 0 ? envio : 0, comision, String(v.cliente || '').slice(0, 80), 0, bruto - comision,
      costoUnit, costoTotal,
      margen, bruto > 0 ? Math.round((margen / bruto) * 10000) / 100 : 0,
      '', '', 'CL',
      new Date(), 'Manual_ERP',
    ];
  });
  const ul = hojaV.getLastRow();
  hojaV.getRange(ul + 1, 1, filas.length, filas[0].length).setValues(filas);
  _formato(hojaV, ul + 1, filas.length);
  return { ok: true, orden: orden, lineas: filas.length, total: filas.reduce((s, f) => s + f[11], 0) };
}

// Solo anula ventas cargadas desde el módulo manual (Origen = Manual_ERP).
function anularManual_(orden, canal) {
  if (!orden || !canal) throw new Error('falta orden o canal');
  const hojaV = SpreadsheetApp.getActive().getSheetByName(H_VENTAS);
  const datos = hojaV.getDataRange().getValues();
  const origenOk = new Set();
  datos.slice(1).forEach(r => { if (String(r[6]) === String(orden) && r[5] === canal && r[25] === 'Manual_ERP') origenOk.add(String(r[6]) + '|' + r[5] + '|' + String(r[7]).trim()); });
  if (!origenOk.size) throw new Error('no se encontró una venta manual ' + orden + ' en ' + canal);
  const n = _anularFilas_(hojaV, k => origenOk.has(k), 'Anulada desde el ERP');
  return { ok: true, anuladas: n };
}

function celdaCsv_(v) {
  v = String(v);
  return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
}

function texto_(t) {
  return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.TEXT);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Genera una clave nueva de 48 caracteres, la guarda y la muestra una vez en el registro.
function generarApiKey() {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Date.now());
  const clave = Utilities.base64EncodeWebSafe(bytes).replace(/[^A-Za-z0-9]/g, '').slice(0, 40) + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  PropertiesService.getScriptProperties().setProperty('API_KEY', clave);
  console.log('Nueva API_KEY (cópiala al secreto SHEETS_API_KEY de Cloudflare): ' + clave);
}
