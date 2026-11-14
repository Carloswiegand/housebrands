// =============================================================
// HOUSEBRANDS — Marketing.gs
// SEO y sitios web de cada marca desde las APIs gratuitas de Google:
//   Search Console (clics, consultas, páginas, indexación), Analytics 4 (sesiones, canales, compras),
//   Merchant Center (estado de productos en Shopping y clics gratuitos) y PageSpeed Insights (velocidad).
// Usa el token de la cuenta dueña del script (ScriptApp.getOAuthToken) con permisos de solo lectura,
// declarados en appsscript.json. No guarda claves.
//
// Hojas:
//   Mkt_Diario     Fecha | Marca | Clics | Impresiones | CTR | Posición | Sesiones | Usuarios |
//                  Sesiones orgánicas | Compras GA4 | Ingresos GA4 | Clics Shopping | Impresiones Shopping
//   Mkt_Consultas  Marca | Consulta | Clics | Impresiones | CTR | Posición | Clics antes | Posición antes
//   Mkt_Paginas    Marca | Página | Clics | Impresiones | CTR | Posición | Clics antes
//   Mkt_Canales    Marca | Canal | Sesiones | Sesiones antes | Compras | Ingresos
//   Mkt_Shopping   Marca | ID | Producto | Estado | Problemas
//   Mkt_Indexacion Marca | URL | Tipo | Veredicto | Estado | Último rastreo | Canónica de Google | Revisado
//   Mkt_Velocidad  Fecha | Marca | URL | Dispositivo | Puntaje | LCP s | INP ms | CLS | Datos reales
// "antes" = los 28 días anteriores a los últimos 28.
// =============================================================

const MKT_MARCAS = [
  { marca: 'MemoKoala', sitio: 'sc-domain:memokoala.com', url: 'https://memokoala.com', ga4: '489510358', mc: '5617286158' },
  { marca: 'Montana', sitio: 'sc-domain:montanagrills.cl', url: 'https://montanagrills.cl', ga4: '528228585', mc: '5745015459' }
];
const MKT_DIAS = 120;          // historia diaria que se reescribe en cada corrida
const MKT_TOP_CONSULTAS = 150;
const MKT_TOP_PAGINAS = 80;
const MKT_MAX_INSPECCION = 250; // URLs por marca en la revisión semanal de indexación (límite Google: 2.000/día)

const MKT_ENC = {
  Mkt_Diario: ['Fecha', 'Marca', 'Clics', 'Impresiones', 'CTR', 'Posición', 'Sesiones', 'Usuarios', 'Sesiones orgánicas', 'Compras GA4', 'Ingresos GA4', 'Clics Shopping', 'Impresiones Shopping'],
  Mkt_Consultas: ['Marca', 'Consulta', 'Clics', 'Impresiones', 'CTR', 'Posición', 'Clics antes', 'Posición antes'],
  Mkt_Paginas: ['Marca', 'Página', 'Clics', 'Impresiones', 'CTR', 'Posición', 'Clics antes'],
  Mkt_Canales: ['Marca', 'Canal', 'Sesiones', 'Sesiones antes', 'Compras', 'Ingresos'],
  Mkt_Shopping: ['Marca', 'ID', 'Producto', 'Estado', 'Problemas'],
  Mkt_Indexacion: ['Marca', 'URL', 'Tipo', 'Veredicto', 'Estado', 'Último rastreo', 'Canónica de Google', 'Revisado'],
  Mkt_Velocidad: ['Fecha', 'Marca', 'URL', 'Dispositivo', 'Puntaje', 'LCP s', 'INP ms', 'CLS', 'Datos reales']
};

// ── utilidades ────────────────────────────────────────────────
function _mkFecha_(diasAtras) { return _smFecha_(new Date(Date.now() - diasAtras * 864e5)); }

function _mkPedir_(url, cuerpo) {
  const op = { method: cuerpo ? 'post' : 'get', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } };
  if (cuerpo) { op.contentType = 'application/json'; op.payload = JSON.stringify(cuerpo); }
  const r = UrlFetchApp.fetch(url, op);
  const txt = r.getContentText();
  let j; try { j = JSON.parse(txt); } catch (e) { j = { raw: txt.slice(0, 200) }; }
  if (r.getResponseCode() >= 300) {
    const m = (j.error && (j.error.message || j.error.status)) || j.raw || ('HTTP ' + r.getResponseCode());
    throw new Error(m);
  }
  return j;
}

// Reemplaza todas las filas de una marca en una hoja (las otras marcas se mantienen)
function _mkReemplazar_(nombre, marca, filas, colMarca) {
  colMarca = colMarca || 0;
  const enc = MKT_ENC[nombre];
  const h = _smHoja_(nombre, enc);
  const resto = _smDatos_(h, enc.length).filter(r => String(r[colMarca]) !== marca);
  const todas = resto.concat(filas);
  if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, enc.length).clearContent();
  if (todas.length) {
    h.getRange(2, 1, todas.length, enc.length).setNumberFormat('@');
    h.getRange(2, 1, todas.length, enc.length).setValues(todas.map(r => r.map(v => v instanceof Date ? _smFecha_(v) : (v === null || v === undefined ? '' : String(v)))));
  }
}

const _mkR_ = (n, d) => { const f = Math.pow(10, d || 0); return Math.round((Number(n) || 0) * f) / f; };

// ── Search Console ────────────────────────────────────────────
function _mkGsc_(sitio, cuerpo) {
  const url = 'https://searchconsole.googleapis.com/webmasters/v3/sites/' + encodeURIComponent(sitio) + '/searchAnalytics/query';
  return (_mkPedir_(url, Object.assign({ dataState: 'all', type: 'web' }, cuerpo)).rows) || [];
}

function _mkTop_(m, dim, n) {
  const ini = _mkFecha_(30), fin = _mkFecha_(3), iniA = _mkFecha_(58), finA = _mkFecha_(31);
  const act = _mkGsc_(m.sitio, { startDate: ini, endDate: fin, dimensions: [dim], rowLimit: n });
  const ant = _mkGsc_(m.sitio, { startDate: iniA, endDate: finA, dimensions: [dim], rowLimit: 1000 });
  const mapa = {}; ant.forEach(r => { mapa[r.keys[0]] = r; });
  return act.map(r => ({ k: r.keys[0], r: r, a: mapa[r.keys[0]] }));
}

// ── Analytics 4 ───────────────────────────────────────────────
function _mkGa4_(prop, cuerpo) {
  return _mkPedir_('https://analyticsdata.googleapis.com/v1beta/properties/' + prop + ':runReport', cuerpo);
}

// ── Merchant Center (Merchant API) ────────────────────────────
function _mkMerchant_(cuenta, consulta) {
  const filas = []; let token = '';
  for (let i = 0; i < 20; i++) {
    const j = _mkPedir_('https://merchantapi.googleapis.com/reports/v1/accounts/' + cuenta + '/reports:search', { query: consulta, pageSize: 1000, pageToken: token || undefined });
    (j.results || []).forEach(x => filas.push(x));
    token = j.nextPageToken; if (!token) break;
  }
  return filas;
}

// ── Carga diaria por marca ────────────────────────────────────
function _mkMarca_(m) {
  const errores = [];
  const ini = _mkFecha_(MKT_DIAS), fin = _mkFecha_(1);
  const dia = {};
  const D = f => (dia[f] = dia[f] || { clics: '', imp: '', ctr: '', pos: '', ses: '', usu: '', org: '', comp: '', ing: '', sc: '', si: '' });

  // Search Console: totales diarios, consultas y páginas
  try {
    _mkGsc_(m.sitio, { startDate: ini, endDate: fin, dimensions: ['date'], rowLimit: 500 }).forEach(r => {
      const d = D(r.keys[0]); d.clics = r.clicks; d.imp = r.impressions; d.ctr = _mkR_(r.ctr * 100, 2); d.pos = _mkR_(r.position, 1);
    });
    _mkReemplazar_('Mkt_Consultas', m.marca, _mkTop_(m, 'query', MKT_TOP_CONSULTAS).map(x =>
      [m.marca, x.k, x.r.clicks, x.r.impressions, _mkR_(x.r.ctr * 100, 2), _mkR_(x.r.position, 1), x.a ? x.a.clicks : 0, x.a ? _mkR_(x.a.position, 1) : '']));
    _mkReemplazar_('Mkt_Paginas', m.marca, _mkTop_(m, 'page', MKT_TOP_PAGINAS).map(x =>
      [m.marca, x.k, x.r.clicks, x.r.impressions, _mkR_(x.r.ctr * 100, 2), _mkR_(x.r.position, 1), x.a ? x.a.clicks : 0]));
  } catch (e) { errores.push('Search Console: ' + e.message); }

  // Analytics 4: diario (con orgánico) y canales 28 días vs 28 anteriores
  try {
    const j = _mkGa4_(m.ga4, {
      dateRanges: [{ startDate: ini, endDate: fin }],
      dimensions: [{ name: 'date' }, { name: 'sessionDefaultChannelGroup' }],
      metrics: [{ name: 'sessions' }, { name: 'totalUsers' }, { name: 'ecommercePurchases' }, { name: 'purchaseRevenue' }],
      limit: 100000
    });
    (j.rows || []).forEach(r => {
      const f = r.dimensionValues[0].value; const fecha = f.slice(0, 4) + '-' + f.slice(4, 6) + '-' + f.slice(6, 8);
      const v = r.metricValues.map(x => Number(x.value) || 0); const d = D(fecha);
      d.ses = (d.ses || 0) + v[0]; d.usu = (d.usu || 0) + v[1]; d.comp = (d.comp || 0) + v[2]; d.ing = (d.ing || 0) + v[3];
      if (r.dimensionValues[1].value === 'Organic Search') d.org = (d.org || 0) + v[0]; else if (d.org === '') d.org = 0;
    });
    const c = _mkGa4_(m.ga4, {
      dateRanges: [{ startDate: _mkFecha_(28), endDate: fin, name: 'act' }, { startDate: _mkFecha_(56), endDate: _mkFecha_(29), name: 'ant' }],
      dimensions: [{ name: 'sessionDefaultChannelGroup' }],
      metrics: [{ name: 'sessions' }, { name: 'ecommercePurchases' }, { name: 'purchaseRevenue' }]
    });
    const can = {};
    (c.rows || []).forEach(r => {
      const nom = r.dimensionValues[0].value; const rango = r.dimensionValues[1] ? r.dimensionValues[1].value : 'act';
      const v = r.metricValues.map(x => Number(x.value) || 0);
      const o = can[nom] = can[nom] || [m.marca, nom, 0, 0, 0, 0];
      if (rango === 'act') { o[2] = v[0]; o[4] = v[1]; o[5] = Math.round(v[2]); } else o[3] = v[0];
    });
    _mkReemplazar_('Mkt_Canales', m.marca, Object.values(can).sort((a, b) => b[2] - a[2]));
  } catch (e) { errores.push('Analytics: ' + e.message); }

  // Merchant Center: estado de productos y clics gratuitos en Shopping
  try {
    const prods = _mkMerchant_(m.mc, 'SELECT id, offer_id, title, aggregated_reporting_context_status, item_issues FROM product_view');
    _mkReemplazar_('Mkt_Shopping', m.marca, prods.map(x => {
      const p = x.productView || {};
      const prob = (p.itemIssues || []).map(i => {
        const t = i.type || {}; const sev = (i.severity && i.severity.aggregatedSeverity) || '';
        return (t.code || '') + (t.canonicalAttribute ? ' (' + t.canonicalAttribute + ')' : '') + (sev ? ' · ' + sev : '');
      }).join(' | ');
      return [m.marca, p.offerId || '', p.title || '', p.aggregatedReportingContextStatus || '', prob.slice(0, 1000)];
    }));
    const perf = _mkMerchant_(m.mc, "SELECT date, clicks, impressions FROM product_performance_view WHERE date BETWEEN '" + ini + "' AND '" + fin + "'");
    perf.forEach(x => {
      const p = x.productPerformanceView || {}; const dd = p.date || {};
      if (!dd.year) return;
      const f = dd.year + '-' + String(dd.month).padStart(2, '0') + '-' + String(dd.day).padStart(2, '0');
      const d = D(f); d.sc = (Number(d.sc) || 0) + (Number(p.clicks) || 0); d.si = (Number(d.si) || 0) + (Number(p.impressions) || 0);
    });
  } catch (e) { errores.push('Merchant Center: ' + e.message); }

  const filas = Object.keys(dia).filter(f => f >= ini && f <= fin).sort().map(f => {
    const d = dia[f];
    return [f, m.marca, d.clics, d.imp, d.ctr, d.pos, d.ses, d.usu, d.org, d.comp, d.ing === '' ? '' : Math.round(d.ing), d.sc, d.si];
  });
  if (filas.length) _mkReemplazar_('Mkt_Diario', m.marca, filas, 1);
  return errores;
}

function marketingActualizar() {
  const err = [];
  MKT_MARCAS.forEach(m => _mkMarca_(m).forEach(e => err.push(m.marca + ' · ' + e)));
  if (err.length) throw new Error(err.join(' / ').slice(0, 900));
}

// ── Revisión semanal: indexación de las URLs del sitemap y velocidad ──
function _mkSitemap_(base) {
  const urls = [];
  const leer = u => { try { const r = UrlFetchApp.fetch(u, { muteHttpExceptions: true }); return r.getResponseCode() === 200 ? r.getContentText() : ''; } catch (e) { return ''; } };
  const idx = leer(base + '/sitemap.xml');
  const hijos = (idx.match(/<loc>[^<]+<\/loc>/g) || []).map(x => x.replace(/<\/?loc>/g, '').replace(/&amp;/g, '&'));
  const mapas = /<sitemapindex/.test(idx) ? hijos : [];
  if (!mapas.length) hijos.forEach(u => urls.push(u));
  mapas.forEach(s => {
    const tipo = /products/.test(s) ? 'Producto' : /collections/.test(s) ? 'Colección' : /blogs/.test(s) ? 'Blog' : /pages/.test(s) ? 'Página' : 'Otro';
    (leer(s).match(/<loc>[^<]+<\/loc>/g) || []).forEach(x => {
      const u = x.replace(/<\/?loc>/g, '').replace(/&amp;/g, '&');
      if (!/\/cdn\/|\.(jpg|png|webp)/i.test(u)) urls.push({ u: u, t: tipo });
    });
  });
  const vistos = {};
  return urls.map(x => typeof x === 'string' ? { u: x, t: 'Otro' } : x).filter(x => !vistos[x.u] && (vistos[x.u] = 1));
}

// fetchAll en paralelo; si Google corta alguna conexión ("Dirección no disponible"), se reintenta una por una
function _mkTodos_(pedidos) {
  try { return UrlFetchApp.fetchAll(pedidos); }
  catch (e) {
    return pedidos.map(p => { try { const u = p.url; const o = Object.assign({}, p); delete o.url; return UrlFetchApp.fetch(u, o); } catch (x) { Utilities.sleep(500); return null; } });
  }
}

function _mkIndexacion_(m) {
  const lista = [{ u: m.url + '/', t: 'Inicio' }].concat(_mkSitemap_(m.url)).slice(0, MKT_MAX_INSPECCION);
  const tok = ScriptApp.getOAuthToken(); const hoy = _smFecha_(); const filas = [];
  for (let i = 0; i < lista.length; i += 20) {
    const grupo = lista.slice(i, i + 20);
    const pedidos = grupo.map(x => ({
      url: 'https://searchconsole.googleapis.com/v1/urlInspection/index:inspect', method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + tok }, payload: JSON.stringify({ inspectionUrl: x.u, siteUrl: m.sitio, languageCode: 'es' })
    }));
    const resp = _mkTodos_(pedidos);
    resp.forEach((r, k) => {
      if (!r) return;
      let j = {}; try { j = JSON.parse(r.getContentText()); } catch (e) { }
      if (r.getResponseCode() >= 300) { if (i === 0 && k === 0) throw new Error((j.error && j.error.message) || 'HTTP ' + r.getResponseCode()); return; }
      const s = (j.inspectionResult && j.inspectionResult.indexStatusResult) || {};
      filas.push([m.marca, grupo[k].u, grupo[k].t, s.verdict || '', s.coverageState || '', (s.lastCrawlTime || '').slice(0, 10), s.googleCanonical || '', hoy]);
    });
    Utilities.sleep(1200);
  }
  _mkReemplazar_('Mkt_Indexacion', m.marca, filas);
}

function _mkVelocidad_(m) {
  const h = _smHoja_('Mkt_Velocidad', MKT_ENC.Mkt_Velocidad);
  // Inicio + la página con más clics orgánicos (si no es el inicio)
  const pags = _smDatos_(_smHoja_('Mkt_Paginas', MKT_ENC.Mkt_Paginas), 3).filter(r => r[0] === m.marca).sort((a, b) => b[2] - a[2]).map(r => r[1]);
  const urls = [m.url + '/'].concat(pags.filter(u => u.replace(/\/$/, '') !== m.url).slice(0, 1));
  const pedidos = [];
  urls.forEach(u => ['mobile', 'desktop'].forEach(s => pedidos.push({ u: u, s: s })));
  const tok = ScriptApp.getOAuthToken();
  const resp = _mkTodos_(pedidos.map(p => ({ url: 'https://www.googleapis.com/pagespeedonline/v5/runPagespeed?category=performance&strategy=' + p.s + '&url=' + encodeURIComponent(p.u), muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + tok } })));
  const hoy = _smFecha_(); const filas = []; let error = '';
  resp.forEach((r, i) => {
    if (!r) { error = error || 'conexión cortada'; return; }
    if (r.getResponseCode() !== 200) { if (!error) { let j = {}; try { j = JSON.parse(r.getContentText()); } catch (e) { } error = (j.error && j.error.message) || 'HTTP ' + r.getResponseCode(); } return; }
    const j = JSON.parse(r.getContentText());
    const lh = j.lighthouseResult || {}; const au = lh.audits || {};
    const ex = (j.loadingExperience && j.loadingExperience.metrics) || {};
    const real = ex.LARGEST_CONTENTFUL_PAINT_MS ? 'Sí' : 'No';
    const lcp = real === 'Sí' ? ex.LARGEST_CONTENTFUL_PAINT_MS.percentile / 1000 : ((au['largest-contentful-paint'] || {}).numericValue || 0) / 1000;
    const inp = ex.INTERACTION_TO_NEXT_PAINT ? ex.INTERACTION_TO_NEXT_PAINT.percentile : '';
    const cls = ex.CUMULATIVE_LAYOUT_SHIFT_SCORE ? ex.CUMULATIVE_LAYOUT_SHIFT_SCORE.percentile / 100 : ((au['cumulative-layout-shift'] || {}).numericValue || 0);
    filas.push([hoy, m.marca, pedidos[i].u, pedidos[i].s === 'mobile' ? 'Móvil' : 'Computador', Math.round(((lh.categories || {}).performance || {}).score * 100) || '', _mkR_(lcp, 1), inp, _mkR_(cls, 2), real]);
  });
  if (filas.length) h.getRange(h.getLastRow() + 1, 1, filas.length, filas[0].length).setNumberFormat('@').setValues(filas.map(r => r.map(String)));
  if (!filas.length) throw new Error(error || 'sin respuesta');
  return filas.length;
}

function marketingSemanal() {
  const err = [];
  MKT_MARCAS.forEach(m => {
    try { _mkIndexacion_(m); } catch (e) { err.push(m.marca + ' · Indexación: ' + e.message); }
    try { _mkVelocidad_(m); } catch (e) { err.push(m.marca + ' · PageSpeed: ' + e.message); }
  });
  if (err.length) throw new Error(err.join(' / ').slice(0, 900));
}

// ── Activadores ───────────────────────────────────────────────
function marketingProgramado() { correr_('Marketing', marketingActualizar); }
function marketingSemanalProgramado() { correr_('Marketing semanal', marketingSemanal); }

// Ejecutar una vez: Marketing todos los días 7:00 y la revisión de indexación y velocidad los lunes 7:30.
function activarMarketing() {
  ScriptApp.getProjectTriggers().filter(t => ['marketingProgramado', 'marketingSemanalProgramado'].indexOf(t.getHandlerFunction()) !== -1).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('marketingProgramado').timeBased().everyDays(1).atHour(7).nearMinute(0).inTimezone('America/Santiago').create();
  ScriptApp.newTrigger('marketingSemanalProgramado').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).nearMinute(30).inTimezone('America/Santiago').create();
  console.log('Marketing: diario 7:00, semanal lunes 7:30.');
}

// Prueba rápida de acceso a cada API (ejecutar desde el editor)
function marketingProbar() {
  const out = [];
  MKT_MARCAS.forEach(m => {
    const t = (n, f) => { try { f(); out.push(m.marca + ' ' + n + ': OK'); } catch (e) { out.push(m.marca + ' ' + n + ': ' + e.message); } };
    t('Search Console', () => _mkGsc_(m.sitio, { startDate: _mkFecha_(10), endDate: _mkFecha_(3), dimensions: ['date'], rowLimit: 5 }));
    t('Analytics', () => _mkGa4_(m.ga4, { dateRanges: [{ startDate: '7daysAgo', endDate: 'yesterday' }], metrics: [{ name: 'sessions' }] }));
    t('Merchant', () => _mkMerchant_(m.mc, 'SELECT id, offer_id FROM product_view'));
  });
  console.log(out.join('\n'));
  return out;
}

// Ejecutar una vez: la Merchant API exige registrar el proyecto de Google Cloud del script en cada cuenta
// de Merchant Center (lo hace un administrador de la cuenta). Esperar 5 minutos antes de usarla.
function marketingRegistrarMerchant() {
  const email = Session.getEffectiveUser().getEmail();
  const out = MKT_MARCAS.map(m => {
    try { _mkPedir_('https://merchantapi.googleapis.com/accounts/v1/accounts/' + m.mc + '/developerRegistration:registerGcp', { developerEmail: email }); return m.marca + ': registrado'; }
    catch (e) { return m.marca + ': ' + e.message; }
  });
  console.log(out.join('\n'));
  return out;
}
