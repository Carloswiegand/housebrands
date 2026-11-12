// =============================================================
// HOUSEBRANDS ERP — Inventario.gs
// Foto diaria del stock publicado en Shopify (MemoKoala + Montana) y MercadoLibre.
// ABEE actualiza el stock de ambas plataformas por su cuenta; aquí solo se LEE
// y se compara. Si Shopify y ML no coinciden para un SKU, queda marcado "Revisar".
//
// Hoja: Inventario_Diario (historial; una fila por SKU por día).
//   Fecha | SKU | Producto | Marca | Stock Shopify | Stock ML | Diferencia | Estado |
//   Publicaciones ML | Stock externo | Stock total
//
// Stock fuera de ABEE (otra bodega, casa HigoKids, etc.) no aparece en las plataformas:
// se anota a mano en la hoja Stock_Externo (SKU | Ubicación | Unidades | Actualizado | Nota)
// y se suma en "Stock total". Stock total = stock ABEE (Shopify; si no hay, ML) + externo.
// =============================================================

const H_INVENTARIO = 'Inventario_Diario';
const H_STOCK_EXT  = 'Stock_Externo';
const ENC_INVENTARIO = ['Fecha', 'SKU', 'Producto', 'Marca', 'Stock Shopify', 'Stock ML', 'Diferencia', 'Estado',
  'Publicaciones ML', 'Stock externo', 'Stock total'];

// Corrida automática diaria (activador). Reemplaza la foto del día si se corre dos veces.
function snapshotInventario() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const sh = _stockShopify_();
    let ml;
    try { ml = _stockML_(); } catch (e) { ml = { stock: {}, nombre: {}, publicaciones: {}, error: e.message }; }
    const ext = _stockExterno_();
    const skus = new Set(Object.keys(COSTOS_SKU).concat(Object.keys(sh.stock), Object.keys(ml.stock), Object.keys(ext)));
    const hoy = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd');

    const filas = [];
    Array.from(skus).filter(Boolean).sort().forEach(sku => {
      const s = sh.stock[sku], m = ml.stock[sku];
      const e = ext[sku] || 0;
      if (s === undefined && m === undefined && !e) return;      // SKU sin publicación ni stock externo (p. ej. extinto)
      const abee = s !== undefined ? s : (m !== undefined ? m : 0);
      const dif = (s === undefined || m === undefined) ? '' : s - m;
      const estado = (s === undefined && m === undefined) ? 'Solo externo' : s === undefined ? 'Solo ML' : m === undefined ? (ml.error ? 'Sin dato ML' : 'Solo Shopify') : dif === 0 ? 'OK' : 'Revisar';
      filas.push([hoy, sku, NOMBRES_SKU[sku] || sh.nombre[sku] || ml.nombre[sku] || '', sh.marca[sku] || _marcaML(sku, ml.nombre[sku] || ''),
        s === undefined ? '' : s, m === undefined ? '' : m, dif, estado, ml.publicaciones[sku] || 0, e, abee + e]);
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName(H_INVENTARIO);
    if (!hoja) { hoja = ss.insertSheet(H_INVENTARIO); hoja.setFrozenRows(1); }
    hoja.getRange(1, 1, 1, ENC_INVENTARIO.length).setValues([ENC_INVENTARIO]);
    // Quita la foto previa del mismo día (si existe) para no duplicar.
    const datos = hoja.getDataRange().getDisplayValues();
    for (let i = datos.length - 1; i >= 1; i--) if (datos[i][0] === hoy) hoja.deleteRow(i + 1);
    if (filas.length) {
      hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
      hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).setNumberFormat('@');
    }

    const revisar = filas.filter(f => ['OK', 'Sin dato ML', 'Solo externo'].indexOf(f[7]) === -1).map(f => f[1] + ' (' + f[7] + ': SH ' + f[4] + ' / ML ' + f[5] + ')');
    const msg = hoy + ' — ' + filas.length + ' SKUs, ' + revisar.length + ' por revisar' + (revisar.length ? ': ' + revisar.join('; ') : '') +
      (sh.errores.length ? ' | errores: ' + sh.errores.join('; ') : '') + (ml.error ? ' | ML: ' + ml.error : '');
    PropertiesService.getScriptProperties().setProperty('INVENTARIO_ULTIMO', msg);
    console.log(msg);
    return msg;
  } finally {
    lock.releaseLock();
  }
}

// Crea (una sola vez) el activador diario a las 23:00 (zona del proyecto).
function activarInventarioDiario() {
  ScriptApp.getProjectTriggers()
    .filter(tr => tr.getHandlerFunction() === 'snapshotInventario')
    .forEach(tr => ScriptApp.deleteTrigger(tr));
  ScriptApp.newTrigger('snapshotInventario').timeBased().everyDays(1).atHour(23).create();
  console.log('Activador diario de snapshotInventario creado (23:00).');
}

// ── internos ────────────────────────────────────────────────

// Unidades fuera de ABEE por SKU (suma de todas las ubicaciones). Crea la hoja si no existe.
function _stockExterno_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(H_STOCK_EXT);
  if (!hoja) {
    hoja = ss.insertSheet(H_STOCK_EXT);
    hoja.appendRow(['SKU', 'Ubicación', 'Unidades', 'Actualizado', 'Nota']);
    hoja.setFrozenRows(1);
    return {};
  }
  const out = {};
  hoja.getDataRange().getValues().slice(1).forEach(r => {
    const sku = String(r[0] || '').trim();
    const u = Number(r[2]) || 0;
    if (sku && u) out[sku] = (out[sku] || 0) + u;
  });
  return out;
}

// Stock disponible por SKU en Shopify (suma de ubicaciones). Si un SKU aparece en
// más de una variante, se toma el mayor (misma existencia física).
function _stockShopify_() {
  const out = { stock: {}, nombre: {}, marca: {}, errores: [] };
  SHOPIFY_TIENDAS.forEach(t => {
    try {
      const token = _shopifyToken_(t);
      if (!token) return;
      const q = 'query($c:String){productVariants(first:100,after:$c){pageInfo{hasNextPage endCursor} ' +
        'nodes{sku inventoryQuantity product{title status vendor}}}}';
      let cursor = null;
      for (let page = 0; page < 20; page++) {
        const r = UrlFetchApp.fetch('https://' + t.shop + '.myshopify.com/admin/api/' + SHOPIFY_API_VERSION + '/graphql.json', {
          method: 'post', contentType: 'application/json',
          headers: { 'X-Shopify-Access-Token': token },
          payload: JSON.stringify({ query: q, variables: { c: cursor } }),
          muteHttpExceptions: true,
        });
        const j = JSON.parse(r.getContentText());
        if (r.getResponseCode() !== 200 || j.errors) throw new Error(JSON.stringify(j.errors || r.getContentText()).slice(0, 200));
        const pv = j.data.productVariants;
        pv.nodes.forEach(n => {
          const sku = String(n.sku || '').trim();
          if (!sku || !n.product || n.product.status === 'ARCHIVED') return;
          const q = Number(n.inventoryQuantity) || 0;
          out.stock[sku] = Math.max(out.stock[sku] === undefined ? -Infinity : out.stock[sku], q);
          out.nombre[sku] = n.product.title;
          out.marca[sku] = t.nombre === 'Montana' ? 'Montana Grills' : 'MemoKoala';
        });
        if (!pv.pageInfo.hasNextPage) break;
        cursor = pv.pageInfo.endCursor;
      }
    } catch (e) {
      out.errores.push(t.nombre + ': ' + e.message);
    }
  });
  return out;
}

// Stock disponible por SKU en MercadoLibre (publicaciones activas y pausadas).
// Varias publicaciones del mismo SKU comparten el stock de ABEE → se toma el mayor.
function _stockML_() {
  const out = { stock: {}, nombre: {}, publicaciones: {} };
  const uid = PropertiesService.getScriptProperties().getProperty('ML_USER_ID');
  const ids = [];
  ['active', 'paused'].forEach(status => {
    for (let offset = 0; offset < 1000; offset += 50) {
      const j = _mlGet_('/users/' + uid + '/items/search?status=' + status + '&limit=50&offset=' + offset);
      (j.results || []).forEach(id => ids.push(id));
      if (!j.results || j.results.length < 50) break;
    }
  });
  const skuDe = x => {
    const a = (x.attributes || []).find(at => at.id === 'SELLER_SKU');
    return String((a && a.value_name) || x.seller_custom_field || '').trim();
  };
  for (let i = 0; i < ids.length; i += 20) {
    const lote = _mlGet_('/items?ids=' + ids.slice(i, i + 20).join(',') + '&attributes=id,title,available_quantity,seller_custom_field,attributes,variations');
    lote.forEach(r => {
      const it = r.body || {};
      const filas = (it.variations && it.variations.length)
        ? it.variations.map(v => ({ sku: skuDe(v), q: v.available_quantity }))
        : [{ sku: skuDe(it), q: it.available_quantity }];
      filas.forEach(f => {
        if (!f.sku) return;
        const q = Number(f.q) || 0;
        out.stock[f.sku] = Math.max(out.stock[f.sku] === undefined ? -Infinity : out.stock[f.sku], q);
        out.nombre[f.sku] = out.nombre[f.sku] || it.title;
        out.publicaciones[f.sku] = (out.publicaciones[f.sku] || 0) + 1;
      });
    });
  }
  return out;
}
