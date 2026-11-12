// =============================================================
// HOUSEBRANDS ERP — ML.gs
// Sincronización automática de ventas MercadoLibre → Libro_Ventas (API ML).
// Reemplaza la carga manual del Excel de "Mis ventas".
//
// Reglas iguales a procesarML():
//   - N° de orden = pack_id (carrito) si existe; si no, id de la orden.
//   - Clave de duplicado: orden|ML|sku (misma función _claves).
//   - Se omiten órdenes canceladas.
//   - Margen = PVP − Costo × IVA (COSTOS_SKU congelados).
//
// Credenciales (Propiedades del script): ML_CLIENT_ID, ML_CLIENT_SECRET.
// Tokens: ML_REFRESH_TOKEN y ML_USER_ID se guardan solos al autorizar.
// =============================================================

const ML_API      = 'https://api.mercadolibre.com';
const ML_AUTH_URL = 'https://auth.mercadolibre.cl/authorization';
const ML_REDIRECT = 'https://script.google.com/macros/s/AKfycbz5zXmoV2Rzd34UtEQNazWnc0SMH9Mopo3Rd3WJt9dSA59bsS1_i92PlfHqmtaP-2sU/exec';
const ML_DIAS_SYNC = 45;   // ventana revisada en cada corrida (nunca antes de ML_DESDE)

// ── Autorización (una sola vez) ─────────────────────────────

// Genera el enlace que debe abrir el dueño de la cuenta de ML para autorizar.
function mlLinkAutorizacion() {
  const state = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('ML_OAUTH_STATE', state);
  const id = PropertiesService.getScriptProperties().getProperty('ML_CLIENT_ID');
  const url = ML_AUTH_URL + '?response_type=code&client_id=' + id +
    '&redirect_uri=' + encodeURIComponent(ML_REDIRECT) + '&state=' + state;
  console.log(url);
  return url;
}

// Lo llama doGet (Api.gs) cuando ML redirige con ?code=...&state=...
function mlCallback_(p) {
  const props = PropertiesService.getScriptProperties();
  const esperado = props.getProperty('ML_OAUTH_STATE');
  if (!esperado || p.state !== esperado) return _html_('No se pudo validar la autorización. Vuelve a generar el enlace.');
  props.deleteProperty('ML_OAUTH_STATE');

  const r = UrlFetchApp.fetch(ML_API + '/oauth/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'authorization_code',
      client_id: props.getProperty('ML_CLIENT_ID'),
      client_secret: props.getProperty('ML_CLIENT_SECRET'),
      code: p.code,
      redirect_uri: ML_REDIRECT,
    },
    muteHttpExceptions: true,
  });
  if (r.getResponseCode() !== 200) return _html_('MercadoLibre rechazó la autorización (HTTP ' + r.getResponseCode() + ').');
  _mlGuardarTokens_(JSON.parse(r.getContentText()));
  return _html_('Listo. MercadoLibre quedó conectado al ERP de HouseBrands. Puedes cerrar esta pestaña.');
}

function _html_(msg) {
  return HtmlService.createHtmlOutput('<div style="font:16px system-ui;padding:40px;max-width:520px">' + msg + '</div>');
}

function _mlGuardarTokens_(j) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ML_REFRESH_TOKEN', j.refresh_token);
  if (j.user_id) props.setProperty('ML_USER_ID', String(j.user_id));
  CacheService.getScriptCache().put('ml_access_token', j.access_token, Math.min(5 * 3600, (j.expires_in || 21600) - 600));
}

// Token de acceso vigente (dura 6 h). El refresh token de ML es de un solo uso: se guarda el nuevo cada vez.
function _mlToken_() {
  const cached = CacheService.getScriptCache().get('ml_access_token');
  if (cached) return cached;
  const props = PropertiesService.getScriptProperties();
  const refresh = props.getProperty('ML_REFRESH_TOKEN');
  if (!refresh) throw new Error('MercadoLibre no está autorizado (falta ML_REFRESH_TOKEN).');
  const r = UrlFetchApp.fetch(ML_API + '/oauth/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'refresh_token',
      client_id: props.getProperty('ML_CLIENT_ID'),
      client_secret: props.getProperty('ML_CLIENT_SECRET'),
      refresh_token: refresh,
    },
    muteHttpExceptions: true,
  });
  if (r.getResponseCode() !== 200) throw new Error('refresh token HTTP ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
  const j = JSON.parse(r.getContentText());
  _mlGuardarTokens_(j);
  return j.access_token;
}

function _mlGet_(path) {
  const r = UrlFetchApp.fetch(ML_API + path, {
    headers: { Authorization: 'Bearer ' + _mlToken_() },
    muteHttpExceptions: true,
  });
  if (r.getResponseCode() !== 200) throw new Error('GET ' + path.split('?')[0] + ' HTTP ' + r.getResponseCode());
  return JSON.parse(r.getContentText());
}

// ── Lectura de órdenes ──────────────────────────────────────

function _mlOrdenes_(desde, hasta) {
  const uid = PropertiesService.getScriptProperties().getProperty('ML_USER_ID');
  const out = [];
  for (let offset = 0; offset < 5000; offset += 50) {
    let q = '/orders/search?seller=' + uid + '&sort=date_asc&limit=50&offset=' + offset +
      '&order.date_created.from=' + encodeURIComponent(desde);
    if (hasta) q += '&order.date_created.to=' + encodeURIComponent(hasta);
    const j = _mlGet_(q);
    (j.results || []).forEach(o => out.push(o));
    if (!j.results || j.results.length < 50) break;
  }
  return out;
}

// Costo de envío y bonificación de un envío según logística:
//   Flex (self_service): envío 0, bonificación = descuento promocionado por ML.
//   Colecta / cross_docking / otros: envío = costo que paga el vendedor, bonificación 0.
function _mlEnvio_(shippingId, cacheEnvios) {
  if (!shippingId) return { costo: 0, bonif: 0 };
  if (cacheEnvios[shippingId]) return cacheEnvios[shippingId];
  let res = { costo: 0, bonif: 0 };
  try {
    const sh = _mlGet_('/shipments/' + shippingId);
    const logistica = String(sh.logistic_type || (sh.logistic || {}).type || '');
    const c = _mlGet_('/shipments/' + shippingId + '/costs');
    const s = (c.senders || [])[0] || {};
    if (logistica === 'self_service') {
      const bonif = (s.discounts || []).reduce((acc, d) => acc + Math.abs(Number(d.promoted_amount) || 0), 0);
      res = { costo: 0, bonif: Math.round(bonif) };
    } else {
      res = { costo: Math.round(Math.abs(Number(s.cost) || 0)), bonif: 0 };
    }
  } catch (e) { /* sin datos de envío: se deja en 0 */ }
  cacheEnvios[shippingId] = res;
  return res;
}

// Convierte órdenes en filas del Libro_Ventas (mismo orden de columnas que procesarML).
function _mlFilas_(ordenes) {
  const filas = [];
  const cacheEnvios = {};
  const envioUsado = {};
  ordenes.forEach(o => {
    if (String(o.status || '').toLowerCase().indexOf('cancel') !== -1) return;
    const num = String(o.pack_id || o.id);
    const fecha = new Date(o.date_created);
    const mes = Number(Utilities.formatDate(fecha, 'America/Santiago', 'M'));
    const anio = Number(Utilities.formatDate(fecha, 'America/Santiago', 'yyyy'));
    const shipId = o.shipping && o.shipping.id;

    (o.order_items || []).forEach(it => {
      const sku = String(it.item.seller_sku || it.item.seller_custom_field || '').trim();
      const titulo = String(it.item.title || '').trim();
      const qty = Number(it.quantity) || 1;
      const precioUnit = Math.round(Number(it.unit_price) || 0);
      const bruto = precioUnit * qty;
      const comision = Math.round((Number(it.sale_fee) || 0) * qty);
      // El envío se asigna una sola vez por envío (un carrito comparte envío).
      let costoEnvio = 0, bonif = 0;
      if (shipId && !envioUsado[shipId]) {
        const e = _mlEnvio_(shipId, cacheEnvios);
        costoEnvio = e.costo; bonif = e.bonif; envioUsado[shipId] = true;
      }
      const ingresoReal = bruto - comision - costoEnvio + bonif;
      const costoUnit = COSTOS_SKU[sku] || 0;
      const costoTotal = costoUnit * qty;
      const margenB = bruto - costoTotal * IVA;
      const margenPct = bruto > 0 ? Math.round((margenB / bruto) * 10000) / 100 : 0;

      filas.push([
        Utilities.formatDate(fecha, 'America/Santiago', 'yyyy-MM-dd'),
        mes, MESES[mes - 1], anio,
        _marcaML(sku, titulo), 'ML', num, sku,
        NOMBRES_SKU[sku] || titulo, qty,
        precioUnit, bruto,
        costoEnvio, comision, '', bonif, ingresoReal,
        costoUnit, costoTotal,
        margenB, margenPct,
        '', '', 'CL',
        new Date(), 'ML_API',
      ]);
    });
  });
  return filas;
}

// ── Comparación (NO escribe nada) ───────────────────────────
// Trae las ventas de ML de un período ya cargado por Excel y compara campo a campo.
function compararML() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const datos = ss.getSheetByName(H_VENTAS).getDataRange().getValues();
  const excel = {};
  let ultimaExcel = '';
  datos.slice(1).forEach(r => {
    if (r[5] !== 'ML') return;
    const f = r[0] instanceof Date ? Utilities.formatDate(r[0], 'America/Santiago', 'yyyy-MM-dd') : String(r[0]);
    if (f > ultimaExcel) ultimaExcel = f;
    excel[String(r[6]) + '|' + String(r[7])] = r;
  });

  const filasApi = _mlFilas_(_mlOrdenes_('2026-09-01T00:00:00.000-03:00', ultimaExcel + 'T23:59:59.000-03:00'));
  const nombres = { 9: 'Cantidad', 11: 'Bruto', 12: 'Envío', 13: 'Comisión', 15: 'Bonificación', 16: 'Ingreso neto' };
  let coinciden = 0, soloApi = [], difs = {};
  Object.keys(nombres).forEach(k => difs[k] = []);
  const vistos = {};
  filasApi.forEach(a => {
    const k = a[6] + '|' + a[7];
    vistos[k] = true;
    const e = excel[k];
    if (!e) { soloApi.push(a[0] + ' ' + k + ' $' + a[11]); return; }
    coinciden++;
    Object.keys(nombres).forEach(c => {
      const va = Number(a[c]) || 0, ve = Number(e[c]) || 0;
      if (Math.abs(va - ve) > 1) difs[c].push(k + ' API=' + va + ' Excel=' + ve);
    });
  });
  const soloExcel = Object.keys(excel).filter(k => {
    const r = excel[k];
    const f = r[0] instanceof Date ? Utilities.formatDate(r[0], 'America/Santiago', 'yyyy-MM-dd') : String(r[0]);
    return f >= '2026-09-01' && !vistos[k];
  });

  const out = [
    'Período: 2026-09-01 a ' + ultimaExcel + ' (última venta ML cargada por Excel)',
    'Líneas API: ' + filasApi.length + ' | coinciden con Excel: ' + coinciden,
    'Solo en API (faltan en el Libro): ' + soloApi.length + (soloApi.length ? ' → ' + soloApi.slice(0, 8).join('; ') : ''),
    'Solo en Excel (no las encontró la API): ' + soloExcel.length + (soloExcel.length ? ' → ' + soloExcel.slice(0, 8).join('; ') : ''),
  ];
  Object.keys(nombres).forEach(c => out.push(nombres[c] + ' distinto: ' + difs[c].length + (difs[c].length ? ' → ' + difs[c].slice(0, 4).join('; ') : '')));
  console.log(out.join('\n'));
  return out;
}

// ── Sincronización (escribe) ────────────────────────────────
// Solo agrega ventas posteriores a la fecha de corte ML_DESDE (yyyy-mm-dd) para no
// mezclar con lo ya cargado por Excel. La deduplicación por orden|ML|sku protege igual.
function syncML() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const corte = props.getProperty('ML_DESDE');
    if (!corte) throw new Error('Falta ML_DESDE (fecha de corte). Se define después de revisar compararML().');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaV = ss.getSheetByName(H_VENTAS);
    const claves = _claves(hojaV);

    const desdeVentana = Utilities.formatDate(new Date(Date.now() - ML_DIAS_SYNC * 864e5), 'America/Santiago', 'yyyy-MM-dd');
    const desde = (desdeVentana > corte ? desdeVentana : corte) + 'T00:00:00.000-03:00';
    const filas = _mlFilas_(_mlOrdenes_(desde)).filter(f => {
      const k = f[6] + '|ML|' + f[7];
      if (claves.has(k)) return false;
      claves.add(k);
      return true;
    });
    if (filas.length) {
      const ul = hojaV.getLastRow();
      hojaV.getRange(ul + 1, 7, filas.length, 1).setNumberFormat('0');
      hojaV.getRange(ul + 1, 1, filas.length, filas[0].length).setValues(filas);
      _formato(hojaV, ul + 1, filas.length);
    }
    const msg = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH:mm') + ' — ' + filas.length + ' líneas nuevas desde ' + desde.slice(0, 10);
    props.setProperty('ML_ULTIMA_SYNC', msg);
    console.log(msg);
    return msg;
  } finally {
    lock.releaseLock();
  }
}

// Muestra los N° de orden de ML completos (sin notación científica) en todo el Libro_Ventas.
function formatoOrdenesML() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_VENTAS);
  h.getRange(2, 7, h.getLastRow() - 1, 1).setNumberFormat('0');
  console.log('Formato de N° de orden aplicado a ' + (h.getLastRow() - 1) + ' filas.');
}

function activarSyncMLHorario() {
  ScriptApp.getProjectTriggers()
    .filter(tr => tr.getHandlerFunction() === 'syncML')
    .forEach(tr => ScriptApp.deleteTrigger(tr));
  ScriptApp.newTrigger('syncML').timeBased().everyHours(1).create();
  console.log('Activador horario de syncML creado.');
}
