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
// POST {key, action:'imp_set', imp:{id,...}}      → crea o reemplaza una importación (hoja Importaciones)
// POST {key, action:'sync'}                       → sincroniza Shopify, ML e inventario en el momento
// POST {key, action:'gastos_set', mes, anio, valores} → gastos variables del mes (hoja Gastos_Variables)
// POST {key, action:'config_set', valores}        → gastos fijos (hoja Config)
// POST {key, action:'pnl_generar'}                → regenera la hoja P&L
// POST {key, action:'gasto_add', doc} / {action:'gasto_del', id} → documentos de gasto (Gastos.gs)
// POST {key, action:'metas_set', filas}           → proyección del flujo (hoja Metas)
// =============================================================

const API_HOJAS = { ventas: 'Libro_Ventas', pnl: 'P&L', conv: 'Conversion', inv: 'Inventario_Diario', imp: 'Importaciones', scout: 'Scout_Tendencias', estado: 'Estado_Sync', gastos: 'Gastos_Variables', config: 'Config', gdet: 'Gastos_Detalle', metas: 'Metas', prefs: 'Preferencias', ideas: 'Scout_Ideas', seg: 'Scout_Seguimiento', opor: 'Scout_Oportunidades', amazon: 'Scout_Amazon', solic: 'Scout_Solicitudes', mktd: 'Mkt_Diario', mktq: 'Mkt_Consultas', mktp: 'Mkt_Paginas', mktc: 'Mkt_Canales', mkts: 'Mkt_Shopping', mkti: 'Mkt_Indexacion', mktv: 'Mkt_Velocidad' };
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
  // Cada sincronización toma su propio candado; por eso va antes del candado general.
  if (body.action === 'sync') {
    try { return json_(sincronizarAhora_()); } catch (err) { return json_({ ok: false, error: err.message }); }
  }
  if (body.action === 'pnl_generar') return json_(correr_('P&L', pnlSinPantalla_));
  if (body.action === 'mkt_sync') { const r = correr_('Marketing', marketingActualizar); return json_({ ok: true, error: r.ok ? '' : r.error }); }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return json_({ ok: false, error: 'ocupado, intenta de nuevo' });
  try {
    if (body.action === 'venta_manual') return json_(ventaManual_(body.venta || {}));
    if (body.action === 'anular_manual') return json_(anularManual_(body.orden, body.canal));
    if (body.action === 'enriquecer') return json_(enriquecerDesdeDatos_(body.canal, body.datos || {}));
    if (body.action === 'imp_set') return json_(importacionGuardar_(body.imp || {}));
    if (body.action === 'gastos_set') return json_(gastosGuardar_(body.mes, body.anio, body.valores || {}));
    if (body.action === 'config_set') return json_(configGuardar_(body.valores || {}));
    if (body.action === 'pref_set') return json_(prefGuardar_(body.clave, body.valor));
    if (body.action === 'scout_idea_add') return json_(scoutIdeaAgregar_(body.idea || {}));
    if (body.action === 'scout_idea_set') return json_(scoutIdeaActualizar_(body.idea || {}));
    if (body.action === 'amazon_add') return json_(amazonAgregar_(body.filas || [], body.fecha));
    if (body.action === 'amazon_clasificar') return json_(amazonClasificar_(body.items || []));
    if (body.action === 'solicitud_add') return json_(solicitudAgregar_(body.solicitud || {}));
    if (body.action === 'solicitud_set') return json_(solicitudActualizar_(body.solicitud || {}));
    if (body.action === 'metas_set') return json_(metasGuardar_(body.filas || []));
    if (body.action === 'gasto_add' || body.action === 'gasto_del') {
      const r = body.action === 'gasto_add' ? gastoDocAgregar_(body.doc || {}) : gastoDocBorrar_(body.id);
      r.pnl = pnlDiferir_();   // el P&L se regenera en segundo plano para no hacer esperar al usuario
      return json_(r);
    }
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

// Completa columnas de análisis desde datos externos (p. ej. CSV histórico de Shopify procesado
// fuera de la planilla). datos = { 'N° orden': [Hora, Cliente, Fuente, Campaña, N° compra, Medio de pago, Despacho, Región código] }.
// Solo llena filas del canal indicado cuya columna Hora está vacía; no toca nada más.
function enriquecerDesdeDatos_(canal, datos) {
  if (['Shopify', 'ML'].indexOf(canal) === -1) throw new Error('canal no válido');
  const hojaV = SpreadsheetApp.getActive().getSheetByName(H_VENTAS);
  _asegurarEncabezadosExtra_(hojaV);
  const n = hojaV.getLastRow() - 1;
  const base = hojaV.getRange(2, 1, n, 26 + H_EXTRA.length).getValues();
  let llenas = 0;
  const extra = base.map(r => {
    const act = r.slice(26, 26 + H_EXTRA.length);
    const d = datos[String(r[6])];
    if (r[5] !== canal || !d || (act[0] !== '' && act[0] !== null)) return act;
    llenas++;
    return H_EXTRA.map((_, j) => (d[j] === undefined || d[j] === null) ? '' : d[j]);
  });
  hojaV.getRange(2, 27, n, H_EXTRA.length).setValues(extra);
  return { ok: true, llenas: llenas };
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
  const truncadas = {};    // ML cargadas por Excel: N° con el último dígito en 0 (Excel guarda 15 dígitos)
  const fechaTxt = v => v instanceof Date ? Utilities.formatDate(v, 'America/Santiago', 'yyyy-MM-dd') : String(v).slice(0, 10);
  base.forEach((r, i) => {
    if (r[26] !== '' && r[26] !== null) return;                       // ya tiene Hora
    if (r[5] !== 'Shopify' && r[5] !== 'ML') return;
    const k = r[5] + '|' + String(r[6]);
    (pendientes[k] = pendientes[k] || []).push(i);
    if (r[5] === 'ML' && /^\d{16}$/.test(String(r[6])) && String(r[6]).slice(-1) === '0') {
      const kt = String(r[6]).slice(0, 15) + '|' + fechaTxt(r[0]);
      (truncadas[kt] = truncadas[kt] || []).push(i);
    }
  });
  let llenas = 0, corregidas = 0;
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
    const num = String(o.pack_id || o.id);
    let k = 'ML|' + num;
    if (!pendientes[k]) {
      // ¿Está cargada con el N° truncado por Excel? Mismo prefijo de 15 dígitos y misma fecha.
      const kt = num.slice(0, 15) + '|' + Utilities.formatDate(new Date(o.date_created), 'America/Santiago', 'yyyy-MM-dd');
      const filas = truncadas[kt];
      if (!filas || num.length !== 16) continue;
      const kViejo = 'ML|' + num.slice(0, 15) + '0';
      filas.forEach(i => { base[i][6] = Number(num); corregidas++; });
      pendientes[k] = (pendientes[kViejo] || []).filter(i => filas.indexOf(i) !== -1);
      if (pendientes[kViejo]) pendientes[kViejo] = pendientes[kViejo].filter(i => filas.indexOf(i) === -1);
      if (pendientes[kViejo] && !pendientes[kViejo].length) delete pendientes[kViejo];
      delete truncadas[kt];
    }
    const env = _mlEnvio_(o.shipping && o.shipping.id, cache);
    const pago = (o.payments || [])[0] || {};
    const medio = String(pago.payment_type || '') + (Number(pago.installments) > 1 ? ' ' + pago.installments + ' cuotas' : '');
    aplicar(k, [Utilities.formatDate(new Date(o.date_created), 'America/Santiago', 'HH:mm'), _clienteId_('ML-', (o.buyer || {}).id),
      'MercadoLibre', '', '', medio, env.logistica, env.regionCod], env.ciudad, env.region);
  }

  // Solo se escriben N° de orden (ML corregidos), Ciudad/Región y las columnas nuevas.
  if (corregidas) hojaV.getRange(2, 7, n, 1).setValues(base.map(r => [r[6]]));
  hojaV.getRange(2, 22, n, 2).setValues(base.map(r => [r[21], r[22]]));
  hojaV.getRange(2, 27, n, H_EXTRA.length).setValues(base.map(r => r.slice(26, 26 + H_EXTRA.length)));
  const msg = llenas + ' líneas completadas, ' + corregidas + ' N° de orden ML corregidos; quedan ' + Object.keys(pendientes).length + ' órdenes sin datos en las APIs' + (cortado ? ' (corte por tiempo: volver a ejecutar)' : '');
  console.log(msg);
  return msg;
}

// ── Importaciones ─────────────────────────────────────────────
// Una fila por importación: ID | Actualizado | JSON con todo el detalle
// (proveedor, pagos, envíos, líneas por SKU, costos estimados y reales, fechas).
function importacionGuardar_(imp) {
  const id = String(imp.id || '').trim();
  if (!/^[A-Za-z0-9._-]{3,60}$/.test(id)) throw new Error('id de importación inválido');
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(API_HOJAS.imp);
  if (!h) { h = ss.insertSheet(API_HOJAS.imp); h.appendRow(['ID', 'Actualizado', 'JSON']); h.setFrozenRows(1); }
  const txt = JSON.stringify(imp);
  if (txt.length > 45000) throw new Error('importación demasiado grande');
  const ids = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const i = ids.indexOf(id);
  const fila = [id, Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH:mm'), txt];
  if (i === -1) h.appendRow(fila); else h.getRange(i + 2, 1, 1, 3).setValues([fila]);
  return { ok: true, id: id, nueva: i === -1 };
}

// ── Sincronización y estado ───────────────────────────────────
// Registra cada corrida en la hoja Estado_Sync (Proceso | Última corrida OK | Último error | Detalle).
function registrarEstado_(proceso, ok, detalle) {
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(API_HOJAS.estado);
  if (!h) { h = ss.insertSheet(API_HOJAS.estado); h.appendRow(['Proceso', 'Última corrida OK', 'Último error', 'Detalle']); h.setFrozenRows(1); }
  const ahora = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH:mm');
  const nombres = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  let i = nombres.indexOf(proceso);
  if (i === -1) { h.appendRow([proceso, '', '', '']); i = nombres.length; }
  const fila = i + 2;
  if (ok) h.getRange(fila, 2).setValue(ahora); else h.getRange(fila, 3).setValue(ahora);
  h.getRange(fila, 4).setValue(String(detalle || '').slice(0, 300));
}

function correr_(proceso, fn) {
  try { fn(); registrarEstado_(proceso, true, 'OK'); return { proceso: proceso, ok: true }; }
  catch (err) { registrarEstado_(proceso, false, err.message); return { proceso: proceso, ok: false, error: err.message }; }
}

// Funciones que llaman los activadores (en vez de llamar directo a syncShopify/syncML/snapshotInventario)
function syncShopifyProgramado() { correr_('Ventas Shopify', syncShopify); }
function syncMLProgramado() { correr_('Ventas MercadoLibre', syncML); }
function inventarioProgramado() { correr_('Inventario', snapshotInventario); }

function sincronizarAhora_() {
  const r = [correr_('Ventas Shopify', syncShopify), correr_('Ventas MercadoLibre', syncML), correr_('Inventario', snapshotInventario), correr_('Conversión', convActualizar)];
  return { ok: r.every(x => x.ok), resultados: r };
}

// Ejecutar una vez: deja Shopify y ML cada 15 minutos e inventario cada hora, con registro de estado.
function activarSincronizacion() {
  const viejas = ['syncShopify', 'syncML', 'snapshotInventario', 'syncShopifyProgramado', 'syncMLProgramado', 'inventarioProgramado'];
  ScriptApp.getProjectTriggers().filter(t => viejas.indexOf(t.getHandlerFunction()) !== -1).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncShopifyProgramado').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('syncMLProgramado').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('inventarioProgramado').timeBased().everyHours(1).create();
  console.log('Sincronización: Shopify y ML cada 15 min, inventario cada hora.');
}

// ── Gastos y P&L desde HB Control ─────────────────────────────
// Variables (cambian cada mes, se cargan con la factura): hoja Gastos_Variables, una fila por mes.
const GASTOS_VARIABLES = ['Meta Ads', 'Google Ads', 'ML Ads', 'ABEE Fulfillment', 'ABEE Envíos', 'The Link It', 'Remuneraciones', 'Otros Admin'];
// Fijos (se repiten cada mes): hoja Config. Los usa PnL.gs.
const GASTOS_FIJOS = ['Shopify MK fee mensual', 'Shopify MT fee mensual', 'Shopify comisión transacciones', 'Claude Pro', 'Google Workspace', 'Canva', 'BP Consulting'];
const MESES_NOMBRE = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

function gastosGuardar_(mes, anio, valores) {
  const m = MESES_NOMBRE.indexOf(String(mes || '').toLowerCase());
  const a = Number(anio);
  if (m === -1 || !(a >= 2024 && a <= 2035)) throw new Error('mes o año inválido');
  const h = SpreadsheetApp.getActive().getSheetByName(API_HOJAS.gastos);
  if (!h) throw new Error('no existe la hoja Gastos_Variables');
  let enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(x => String(x).trim());
  Object.keys(valores).forEach(k => {
    if (GASTOS_VARIABLES.indexOf(k) === -1) throw new Error('concepto no permitido: ' + k);
    if (enc.indexOf(k) === -1) { h.getRange(1, enc.length + 1).setValue(k).setFontWeight('bold'); enc.push(k); }
  });
  const datos = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, 3).getValues() : [];
  let fila = datos.findIndex(r => String(r[1]).trim().toLowerCase() === MESES_NOMBRE[m] && Number(r[2]) === a);
  if (fila === -1) { fila = datos.length; h.getRange(fila + 2, 1, 1, 3).setValues([[m + 1, MESES_NOMBRE[m].charAt(0).toUpperCase() + MESES_NOMBRE[m].slice(1), a]]); }
  Object.keys(valores).forEach(k => {
    const v = valores[k];
    h.getRange(fila + 2, enc.indexOf(k) + 1).setValue(v === '' || v === null ? '' : Math.round(Number(v) || 0));
  });
  return { ok: true, mes: MESES_NOMBRE[m], anio: a };
}

function configGuardar_(valores) {
  const h = SpreadsheetApp.getActive().getSheetByName(API_HOJAS.config);
  if (!h) throw new Error('no existe la hoja Config');
  const datos = h.getDataRange().getValues();
  Object.keys(valores).forEach(k => {
    if (GASTOS_FIJOS.indexOf(k) === -1) throw new Error('concepto no permitido: ' + k);
    const v = Number(valores[k]);
    if (!isFinite(v) || v < 0) throw new Error('valor inválido en ' + k);
    const i = datos.findIndex(r => String(r[0]).trim() === k);
    if (i === -1) { h.appendRow([k, v]); datos.push([k, v]); } else h.getRange(i + 1, 2).setValue(v);
  });
  return { ok: true };
}

// generarPnL() muestra avisos en pantalla; aquí corre sin pantalla y deja pasar los errores.
var PNL_SIN_PANTALLA = false;
function pnlSinPantalla_() { PNL_SIN_PANTALLA = true; try { generarPnL(); } finally { PNL_SIN_PANTALLA = false; } }
function pnlProgramado() { correr_('P&L', pnlSinPantalla_); }

// P&L en segundo plano: un activador de una sola vez que corre ~15 s después de guardar.
// Si se guardan varios gastos seguidos, queda uno solo pendiente.
function pnlDiferir_() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'pnlDiferido').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pnlDiferido').timeBased().after(15 * 1000).create();
  return 'programado';
}
function pnlDiferido() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'pnlDiferido').forEach(t => ScriptApp.deleteTrigger(t));
  correr_('P&L', pnlSinPantalla_);
}

// Ejecutar una vez: regenera el P&L todas las noches (el último día del mes incluye los gastos fijos).
function activarPnLDiario() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'pnlProgramado').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('pnlProgramado').timeBased().everyDays(1).atHour(23).nearMinute(30).create();
  console.log('P&L diario activado (23:30).');
}

// ── Preferencias de HB Control ────────────────────────────────
// Hoja Preferencias: Clave | Valor (JSON) | Actualizado. Todo lo que antes quedaba solo en el navegador
// (lista de cierre mensual, plazos de importación, evaluador del Scout) se guarda acá.
function prefGuardar_(clave, valor) {
  clave = String(clave || '').trim();
  if (!/^[A-Za-z0-9_.-]{1,60}$/.test(clave)) throw new Error('clave inválida');
  const txt = JSON.stringify(valor === undefined ? null : valor);
  if (txt.length > 45000) throw new Error('valor demasiado grande');
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(API_HOJAS.prefs);
  if (!h) { h = ss.insertSheet(API_HOJAS.prefs); h.appendRow(['Clave', 'Valor', 'Actualizado']); h.setFrozenRows(1); h.getRange(1, 1, 1, 3).setFontWeight('bold'); }
  const claves = h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const i = claves.indexOf(clave);
  const fila = [clave, txt, Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH:mm')];
  if (i === -1) h.appendRow(fila); else h.getRange(i + 2, 1, 1, 3).setValues([fila]);
  return { ok: true };
}
