// =============================================================
// HOUSEBRANDS ERP — Shopify.gs
// Sincronización automática de pedidos Shopify → Libro_Ventas (Admin API).
// Reemplaza la importación manual de CSV de Shopify.
//
// Reutiliza las mismas reglas de procesarShopify(): COSTOS_SKU, NOMBRES_SKU,
// IVA, TASA_MP, _marcaSH, _claves (clave orden|Shopify|sku) y _formato.
// Una venta importada por API queda idéntica a una importada por CSV,
// salvo la columna Origen = 'Shopify_API'.
//
// Credenciales por tienda (Propiedades del script, nunca en el código):
//   SHOPIFY_MK_CLIENT_ID / SHOPIFY_MK_CLIENT_SECRET   → MemoKoala
//   SHOPIFY_MG_CLIENT_ID / SHOPIFY_MG_CLIENT_SECRET   → Montana
// Si faltan las de una tienda, esa tienda se omite.
// =============================================================

const SHOPIFY_API_VERSION = '2026-07';
const SHOPIFY_DIAS_SYNC   = 45;   // ventana que se revisa en cada corrida (la deduplicación evita repetidos)
const SHOPIFY_TIENDAS = [
  { key: 'MK', shop: 'fe487a-33',    nombre: 'MemoKoala' },
  { key: 'MG', shop: 'montana-9727', nombre: 'Montana'   },
];

// Corrida automática (la llama el activador horario). No muestra alertas.
function syncShopify() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const res = SHOPIFY_TIENDAS.map(t => _syncTienda_(t));
    const txt = res.map(r => r.tienda + ': ' + r.msg).join(' | ');
    PropertiesService.getScriptProperties().setProperty('SHOPIFY_ULTIMA_SYNC',
      Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH:mm') + ' — ' + txt);
    console.log(txt);
    return res;
  } finally {
    lock.releaseLock();
  }
}

// Para ejecutar a mano desde el menú: muestra el resultado.
function syncShopifyManual() {
  const res = syncShopify() || [];
  SpreadsheetApp.getUi().alert('Sincronización Shopify', res.map(r => r.tienda + ': ' + r.msg).join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// Prueba de conexión: no escribe nada, solo cuenta pedidos recientes por tienda.
function probarShopify() {
  const out = SHOPIFY_TIENDAS.map(t => {
    try {
      const token = _shopifyToken_(t);
      if (!token) return t.nombre + ': sin credenciales';
      const pedidos = _shopifyPedidos_(t, token, 7);
      return t.nombre + ': OK, ' + pedidos.length + ' pedidos en 7 días' + (pedidos[0] ? ' (último ' + pedidos[pedidos.length - 1].name + ')' : '');
    } catch (e) {
      return t.nombre + ': ERROR ' + e.message;
    }
  });
  console.log(out.join('\n'));
  return out;
}

// Crea (una sola vez) el activador que corre syncShopify cada hora.
function activarSyncShopifyHorario() {
  ScriptApp.getProjectTriggers()
    .filter(tr => tr.getHandlerFunction() === 'syncShopify')
    .forEach(tr => ScriptApp.deleteTrigger(tr));
  ScriptApp.newTrigger('syncShopify').timeBased().everyHours(1).create();
  console.log('Activador horario de syncShopify creado.');
}

// ── internos ────────────────────────────────────────────────

function _syncTienda_(t) {
  try {
    const token = _shopifyToken_(t);
    if (!token) return { tienda: t.nombre, msg: 'omitida (sin credenciales)' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hojaV = ss.getSheetByName(H_VENTAS);
    if (!hojaV) { hojaV = ss.insertSheet(H_VENTAS); _encabezados(hojaV); }
    const pedidos = _shopifyPedidos_(t, token, SHOPIFY_DIAS_SYNC);

    // Pedidos cancelados o reembolsados completos después de registrados → Ventas_Anuladas.
    const anulados = new Set();
    pedidos.forEach(o => {
      const fin = String(o.displayFinancialStatus || '').toLowerCase();
      if (o.cancelledAt || fin === 'refunded' || fin === 'voided') {
        (o.lineItems.nodes || []).forEach(li => anulados.add(o.name + '|Shopify|' + String(li.sku || '').trim()));
      }
    });
    const nAnul = _anularFilas_(hojaV, k => anulados.has(k), 'Cancelado/reembolsado en Shopify ' + t.nombre);
    const claves = _claves(hojaV);

    const rows = [];
    let dup = 0, omit = 0;

    pedidos.forEach(o => {
      const fin = String(o.displayFinancialStatus || '').toLowerCase();
      if (o.test || o.cancelledAt || fin === 'refunded' || fin === 'voided') { omit++; return; }

      const fecha = new Date(o.processedAt || o.createdAt);
      const mes = Number(Utilities.formatDate(fecha, 'America/Santiago', 'M'));
      const anio = Number(Utilities.formatDate(fecha, 'America/Santiago', 'yyyy'));
      const envioTotal = _monto_(o.totalShippingPriceSet);
      const dir = o.billingAddress || o.shippingAddress || {};

      (o.lineItems.nodes || []).forEach((li, idx) => {
        const sku = String(li.sku || '').trim();
        const itemNom = String(li.name || '').trim();
        const clave = o.name + '|Shopify|' + sku;
        if (claves.has(clave)) { dup++; return; }

        const qty = Number(li.quantity) || 1;
        const precioUnit = _monto_(li.originalUnitPriceSet);
        const precioTotal = precioUnit * qty;
        const descuento = _monto_(li.totalDiscountSet);
        const precioBruto = precioTotal - descuento;
        const costoUnit = COSTOS_SKU[sku] || 0;
        const costoTotal = costoUnit * qty;
        const margenB = precioBruto - costoTotal * IVA;
        const margenPct = precioBruto > 0 ? Math.round((margenB / precioBruto) * 10000) / 100 : 0;
        const cargoMP = Math.round(precioBruto * TASA_MP);
        // Igual que el CSV: envío y código de descuento solo en la primera línea del pedido.
        const envio = idx === 0 ? envioTotal : 0;
        const discCode = idx === 0 ? String(o.discountCode || '') : '';
        const marca = _marcaSH(String(li.vendor || '').toLowerCase(), sku, itemNom);

        rows.push([
          Utilities.formatDate(fecha, 'America/Santiago', 'yyyy-MM-dd'),
          mes, MESES[mes - 1], anio,
          marca, 'Shopify', o.name, sku,
          NOMBRES_SKU[sku] || itemNom, qty,
          precioUnit, precioBruto,
          envio, cargoMP, discCode, descuento, precioBruto - cargoMP,
          costoUnit, costoTotal,
          margenB, margenPct,
          String(dir.city || ''), String(dir.province || ''), String(dir.countryCodeV2 || ''),
          new Date(), 'Shopify_API',
        ]);
        claves.add(clave);
      });
    });

    if (rows.length > 0) {
      _avisarSinCosto_(rows, 'Shopify ' + t.nombre);
      const ul = hojaV.getLastRow();
      hojaV.getRange(ul + 1, 1, rows.length, rows[0].length).setValues(rows);
      _formato(hojaV, ul + 1, rows.length);
    }
    return { tienda: t.nombre, msg: rows.length + ' líneas nuevas, ' + dup + ' ya existían, ' + omit + ' pedidos omitidos, ' + nAnul + ' líneas anuladas' };
  } catch (e) {
    return { tienda: t.nombre, msg: 'ERROR ' + e.message };
  }
}

function _shopifyToken_(t) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHOPIFY_' + t.key + '_CLIENT_ID');
  const secret = props.getProperty('SHOPIFY_' + t.key + '_CLIENT_SECRET');
  if (!id || !secret) return null;

  // Los tokens duran 24 h: se guardan en caché 6 h (máximo de CacheService).
  const cache = CacheService.getScriptCache();
  const ck = 'shopify_token_' + t.key;
  const cached = cache.get(ck);
  if (cached) return cached;

  const r = UrlFetchApp.fetch('https://' + t.shop + '.myshopify.com/admin/oauth/access_token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: { grant_type: 'client_credentials', client_id: id, client_secret: secret },
    muteHttpExceptions: true,
  });
  if (r.getResponseCode() !== 200) throw new Error('token HTTP ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
  const token = JSON.parse(r.getContentText()).access_token;
  cache.put(ck, token, 6 * 60 * 60); // máximo permitido por CacheService (6 h)
  return token;
}

function _shopifyPedidos_(t, token, dias) {
  const desde = Utilities.formatDate(new Date(Date.now() - dias * 864e5), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const query = 'query($c:String,$q:String){orders(first:100,after:$c,query:$q,sortKey:CREATED_AT){' +
    'pageInfo{hasNextPage endCursor} nodes{name createdAt processedAt cancelledAt test displayFinancialStatus discountCode ' +
    'totalShippingPriceSet{shopMoney{amount}} billingAddress{city province countryCodeV2} shippingAddress{city province countryCodeV2} ' +
    'lineItems(first:50){nodes{sku name vendor quantity originalUnitPriceSet{shopMoney{amount}} totalDiscountSet{shopMoney{amount}}}}}}}';
  const out = [];
  let cursor = null;
  for (let page = 0; page < 50; page++) {
    const r = UrlFetchApp.fetch('https://' + t.shop + '.myshopify.com/admin/api/' + SHOPIFY_API_VERSION + '/graphql.json', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'X-Shopify-Access-Token': token },
      payload: JSON.stringify({ query: query, variables: { c: cursor, q: 'created_at:>=' + desde } }),
      muteHttpExceptions: true,
    });
    if (r.getResponseCode() !== 200) throw new Error('GraphQL HTTP ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
    const j = JSON.parse(r.getContentText());
    if (j.errors) throw new Error(JSON.stringify(j.errors).slice(0, 300));
    const ord = j.data.orders;
    ord.nodes.forEach(n => out.push(n));
    if (!ord.pageInfo.hasNextPage) break;
    cursor = ord.pageInfo.endCursor;
  }
  return out;
}

function _monto_(set) {
  return set && set.shopMoney ? Math.round(Number(set.shopMoney.amount) || 0) : 0;
}
