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
      '', _clienteId_('MN-', v.cliente), canal, '', '', '', '', '',
    ];
  });
  _asegurarEncabezadosExtra_(hojaV);
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

// ── Aviso de SKU sin costo ──────────────────────────────────
// Las ventas automáticas (Shopify/ML) se registran igual, pero con costo 0 el margen sale
// inflado. Se avisa por correo al dueño del script una vez por SKU al día.
function _avisarSinCosto_(filas, canal) {
  const sinCosto = {};
  filas.forEach(f => { if (!COSTOS_SKU[f[7]]) (sinCosto[f[7] || '(sin SKU)'] = sinCosto[f[7] || '(sin SKU)'] || []).push(f[6]); });
  const skus = Object.keys(sinCosto);
  if (!skus.length) return;
  const cache = CacheService.getScriptCache();
  const nuevos = skus.filter(s => !cache.get('aviso_sin_costo_' + s));
  if (!nuevos.length) return;
  const cuerpo = 'Entraron ventas de ' + canal + ' con SKU que no tiene costo cargado en el ERP.\n' +
    'Quedaron registradas con costo $0, así que su margen aparece inflado hasta que se cargue el costo.\n\n' +
    nuevos.map(s => '• ' + s + ' — órdenes: ' + sinCosto[s].join(', ')).join('\n') +
    '\n\nQué hacer: agregar el SKU a COSTOS_SKU y NOMBRES_SKU (Código.gs) y luego recalcular márgenes.';
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'ERP HouseBrands: venta con SKU sin costo (' + nuevos.join(', ') + ')', cuerpo);
    nuevos.forEach(s => cache.put('aviso_sin_costo_' + s, '1', 6 * 3600));
  } catch (e) { console.log('No se pudo enviar el aviso: ' + e.message); }
}

// Ejecutar una vez desde el editor para autorizar el envío de correos.
function probarAvisoSinCosto() {
  ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, ['https://www.googleapis.com/auth/script.send_mail']);
  _avisarSinCosto_([[0, 0, 0, 0, 0, 'Prueba', 'PRUEBA-1', 'SKU-DE-PRUEBA']], 'Prueba');
}

// ── Columnas de análisis (después de Origen) ────────────────
// Se agregan al final para no mover las 26 columnas que usan el P&L y el ERP.
const H_EXTRA = ['Hora', 'Cliente', 'Fuente', 'Campaña', 'N° compra cliente', 'Medio de pago', 'Despacho', 'Región código'];

function _asegurarEncabezadosExtra_(hojaV) {
  const r = hojaV.getRange(1, 27, 1, H_EXTRA.length);
  const act = r.getValues()[0];
  if (act.join('|') !== H_EXTRA.join('|')) r.setValues([H_EXTRA]).setFontWeight('bold');
}

// Código anónimo y estable de cliente (no se guarda el email ni el id real).
function _clienteId_(prefijo, valor) {
  const v = String(valor || '').trim().toLowerCase();
  if (!v) return '';
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'hb|' + v);
  return prefijo + d.slice(0, 5).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

// Código de región ISO de Chile sin prefijo ("CL-RM" → "RM").
function _regionCod_(v) { return String(v || '').replace(/^CL-/, '').trim().toUpperCase(); }

// Rellena columnas de análisis (y ciudad/región de ML) en ventas ya registradas, con lo que
// las APIs todavía entregan: Shopify últimos ~60 días, ML último año. Se puede correr varias
// veces: solo completa filas vacías y se detiene antes del límite de tiempo.
function enriquecerHistorico() {
  const t0 = Date.now();
  const hojaV = SpreadsheetApp.getActive().getSheetByName(H_VENTAS);
  _asegurarEncabezadosExtra_(hojaV);
  const n = hojaV.getLastRow() - 1;
  const base = hojaV.getRange(2, 1, n, 26 + H_EXTRA.length).getValues();
  const pendientes = {};   // 'canal|orden' → [índices de fila]
  base.forEach((r, i) => {
    if (r[26] !== '' && r[26] !== null) return;                       // ya tiene Hora
    if (r[5] !== 'Shopify' && r[5] !== 'ML') return;
    const k = r[5] + '|' + String(r[6]);
    (pendientes[k] = pendientes[k] || []).push(i);
  });
  let llenas = 0;
  const aplicar = (k, extra, ciudad, region) => {
    (pendientes[k] || []).forEach(i => {
      for (let j = 0; j < H_EXTRA.length; j++) base[i][26 + j] = extra[j];
      if (ciudad && !base[i][21]) base[i][21] = ciudad;
      if (region && !base[i][22]) base[i][22] = region;
      llenas++;
    });
    delete pendientes[k];
  };

  // Shopify
  SHOPIFY_TIENDAS.forEach(t => {
    const token = _shopifyToken_(t);
    if (!token) return;
    _shopifyPedidos_(t, token, 60).forEach(o => {
      const k = 'Shopify|' + o.name;
      if (pendientes[k]) aplicar(k, _shopifyExtra_(o, new Date(o.processedAt || o.createdAt)));
    });
  });

  // MercadoLibre (último año). Solo consulta el envío de órdenes que faltan.
  const desde = Utilities.formatDate(new Date(Date.now() - 365 * 864e5), 'America/Santiago', "yyyy-MM-dd'T'00:00:00.000-03:00");
  const cache = {};
  let cortado = false;
  for (const o of _mlOrdenes_(desde)) {
    if (Date.now() - t0 > 270000) { cortado = true; break; }
    const k = 'ML|' + String(o.pack_id || o.id);
    if (!pendientes[k]) continue;
    const env = _mlEnvio_(o.shipping && o.shipping.id, cache);
    const pago = (o.payments || [])[0] || {};
    const medio = String(pago.payment_type || '') + (Number(pago.installments) > 1 ? ' ' + pago.installments + ' cuotas' : '');
    aplicar(k, [Utilities.formatDate(new Date(o.date_created), 'America/Santiago', 'HH:mm'), _clienteId_('ML-', (o.buyer || {}).id),
      'MercadoLibre', '', '', medio, env.logistica, env.regionCod], env.ciudad, env.region);
  }

  // Solo se escriben Ciudad/Región y las columnas nuevas (el resto de la fila no se toca).
  hojaV.getRange(2, 22, n, 2).setValues(base.map(r => [r[21], r[22]]));
  hojaV.getRange(2, 27, n, H_EXTRA.length).setValues(base.map(r => r.slice(26, 26 + H_EXTRA.length)));
  const msg = llenas + ' líneas completadas; quedan ' + Object.keys(pendientes).length + ' órdenes sin datos en las APIs' + (cortado ? ' (corte por tiempo: volver a ejecutar)' : '');
  console.log(msg);
  return msg;
}
