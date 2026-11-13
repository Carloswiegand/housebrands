// =============================================================
// HOUSEBRANDS ERP — Conversion.gs
// Llena la hoja Conversion sola, para el mes en curso y el anterior:
//   Mes | Año | Canal | Sesiones | Carro | Checkout | Compras
// - Shopify (cada tienda): sesiones, sesiones con carro, sesiones que llegaron al checkout y
//   sesiones que compraron, desde ShopifyQL (necesita el permiso read_reports en la app).
// - MercadoLibre: visitas a las publicaciones (API de visitas) y ventas del Libro_Ventas.
// Corre cada hora (las cifras de Shopify Analytics se consolidan con algo de retraso).
// =============================================================

const CONV_MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const CONV_CANAL_SH = { MK: 'Shopify MK', MG: 'Shopify MT' };

function convActualizar() {
  const hoy = new Date();
  const meses = [new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1), new Date(hoy.getFullYear(), hoy.getMonth(), 1)];
  const filas = [], errores = [];
  meses.forEach(ini => {
    const fin = new Date(ini.getFullYear(), ini.getMonth() + 1, 0);
    const hasta = fin > hoy ? hoy : fin;
    const mes = CONV_MESES[ini.getMonth()], anio = ini.getFullYear();
    SHOPIFY_TIENDAS.forEach(t => {
      try {
        const m = _convShopify_(t, ini, hasta);
        if (m) filas.push([mes, anio, CONV_CANAL_SH[t.key] || ('Shopify ' + t.key), m.sesiones, m.carro, m.checkout, m.compras]);
      } catch (e) { errores.push(t.nombre + ': ' + e.message); }
    });
    try {
      const v = _convMLVisitas_(ini, hasta);
      filas.push([mes, anio, 'ML', v, '', '', _convMLVentas_(ini, hasta)]);
    } catch (e) { errores.push('ML: ' + e.message); }
  });
  _convEscribir_(filas);
  if (errores.length) throw new Error(errores.join(' | ').slice(0, 280));
  return filas.length;
}

function convProgramado() { correr_('Conversión', convActualizar); }

// Ejecutar una vez: deja la conversión actualizándose cada hora.
function activarConversion() {
  ScriptApp.getProjectTriggers().filter(t => ['convProgramado', 'convActualizar'].indexOf(t.getHandlerFunction()) !== -1).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('convProgramado').timeBased().everyHours(1).create();
  console.log('Conversión automática activada (cada hora).');
}

function _convFecha_(d) { return Utilities.formatDate(d, 'America/Santiago', 'yyyy-MM-dd'); }

function _convShopify_(t, desde, hasta) {
  const token = _shopifyToken_(t);
  if (!token) return null;
  const q = 'FROM sessions SHOW sessions, sessions_with_cart_additions, sessions_that_reached_checkout, sessions_that_completed_checkout ' +
    'SINCE ' + _convFecha_(desde) + ' UNTIL ' + _convFecha_(hasta);
  const body = { query: 'query($q:String!){shopifyqlQuery(query:$q){tableData{columns{name} rows} parseErrors}}', variables: { q: q } };
  const r = UrlFetchApp.fetch('https://' + t.shop + '.myshopify.com/admin/api/' + SHOPIFY_API_VERSION + '/graphql.json', {
    method: 'post', contentType: 'application/json', headers: { 'X-Shopify-Access-Token': token },
    payload: JSON.stringify(body), muteHttpExceptions: true,
  });
  if (r.getResponseCode() !== 200) throw new Error('ShopifyQL HTTP ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 160));
  const j = JSON.parse(r.getContentText());
  if (j.errors) throw new Error(JSON.stringify(j.errors).slice(0, 200));
  const res = j.data && j.data.shopifyqlQuery;
  if (!res) throw new Error('sin respuesta de ShopifyQL');
  if (res.parseErrors && res.parseErrors.length) throw new Error('ShopifyQL: ' + JSON.stringify(res.parseErrors).slice(0, 200));
  const cols = (res.tableData.columns || []).map(c => c.name);
  const rows = res.tableData.rows || [];
  const fila = rows[0];
  const val = n => {
    const i = cols.indexOf(n);
    if (i === -1 || !fila) return 0;
    const v = Array.isArray(fila) ? fila[i] : fila[n];
    return Number(v) || 0;
  };
  return { sesiones: val('sessions'), carro: val('sessions_with_cart_additions'), checkout: val('sessions_that_reached_checkout'), compras: val('sessions_that_completed_checkout') };
}

function _convMLVisitas_(desde, hasta) {
  const uid = PropertiesService.getScriptProperties().getProperty('ML_USER_ID');
  const fin = new Date(hasta.getFullYear(), hasta.getMonth(), hasta.getDate() + 1);
  const j = _mlGet_('/users/' + uid + '/items_visits?date_from=' + _convFecha_(desde) + 'T00:00:00.000-03:00&date_to=' + _convFecha_(fin) + 'T00:00:00.000-03:00');
  return Number(j.total_visits) || 0;
}

// Pedidos únicos de ML en el Libro_Ventas dentro del rango.
function _convMLVentas_(desde, hasta) {
  const h = SpreadsheetApp.getActive().getSheetByName(H_VENTAS);
  const d = h.getRange(2, 1, h.getLastRow() - 1, 7).getValues();
  const a = _convFecha_(desde), b = _convFecha_(hasta);
  const ord = new Set();
  d.forEach(r => {
    if (r[5] !== 'ML') return;
    const f = r[0] instanceof Date ? _convFecha_(r[0]) : String(r[0]).slice(0, 10);
    if (f >= a && f <= b) ord.add(String(r[6]));
  });
  return ord.size;
}

// Reemplaza las filas del mismo mes, año y canal; agrega las que no existen.
function _convEscribir_(filas) {
  if (!filas.length) return;
  const h = SpreadsheetApp.getActive().getSheetByName('Conversion');
  if (!h) throw new Error('no existe la hoja Conversion');
  const n = h.getLastRow();
  const datos = n > 1 ? h.getRange(2, 1, n - 1, 7).getValues() : [];
  const clave = r => String(r[0]).trim().toLowerCase() + '|' + Number(r[1]) + '|' + String(r[2]).trim();
  const idx = {};
  datos.forEach((r, i) => { idx[clave(r)] = i; });
  filas.forEach(f => {
    const i = idx[clave(f)];
    if (i === undefined) { h.appendRow(f); }
    else h.getRange(i + 2, 1, 1, 7).setValues([f]);
  });
}
