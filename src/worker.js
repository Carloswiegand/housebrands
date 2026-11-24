// HouseBrands — Worker de housebrands.cl
// Sirve el sitio estático (carpeta /public) y expone la API privada del ERP:
//   /erp/api/ventas  → hoja Libro_Ventas
//   /erp/api/pnl     → hoja P&L
//   /erp/api/conv    → hoja Conversion
//   /erp/api/inv     → hoja Inventario_Diario
//   POST /erp/api/venta → registra o anula una venta manual (Apps Script doPost)
// La API solo responde si la petición trae un token válido de Cloudflare Access
// emitido para un correo autorizado (ALLOWED_EMAILS). Falla cerrada: si falta
// cualquier configuración, responde 503 y no entrega datos.
//
// Configuración (wrangler.jsonc → vars):  ACCESS_TEAM_DOMAIN, ACCESS_AUD, ALLOWED_EMAILS, SHEETS_API_URL
// Secreto (dashboard / wrangler secret):   SHEETS_API_KEY
// Opcional: SCOUT_ROUTINE_ID (var) + SCOUT_FIRE_TOKEN (secreto) → al pedir un análisis del Scout se dispara
// en el momento la tarea programada "Scout de productos: Amazon y análisis" (API de rutinas de Claude Code).

const SHEETS = new Set(['ventas', 'pnl', 'conv', 'inv', 'imp', 'scout', 'estado', 'gastos', 'config', 'gdet', 'metas', 'prefs', 'ideas', 'seg', 'opor', 'amazon', 'solic', 'mktd', 'mktq', 'mktp', 'mktc', 'mkts', 'mkti', 'mktv', 'bench', 'contab']);
const CACHE_SECONDS = 60;

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (url.pathname === '/api/pedidos') return pedidosPublicos(env, ctx);
    if (url.pathname === '/api/precios-web') return preciosWeb(ctx);
    const m = url.pathname.match(/^\/erp\/api\/([a-z]+)\/?$/);
    if (m && m[1] === 'venta') return ventaManual(request, env);
    if (m) return api(m[1], request, env, ctx);
    return env.ASSETS.fetch(request);
  },
};

async function api(sheet, request, env, ctx) {
  if (!SHEETS.has(sheet)) return json({ error: 'not_found' }, 404);
  if (!env.ACCESS_TEAM_DOMAIN || !env.ACCESS_AUD || !env.ALLOWED_EMAILS || !env.SHEETS_API_URL || !env.SHEETS_API_KEY) {
    return json({ error: 'not_configured' }, 503);
  }

  const identity = await verifyAccess(request, env).catch(() => null);
  if (!identity) return json({ error: 'unauthorized' }, 401);

  // Caché interna corta (después de validar identidad; nunca se sirve sin autenticación).
  const cacheKey = new Request(`https://cache.internal/erp/${sheet}`);
  const cache = caches.default;
  const hit = await cache.match(cacheKey);
  if (hit) return withPrivateHeaders(hit);

  const upstream = new URL(env.SHEETS_API_URL);
  upstream.searchParams.set('sheet', sheet);
  upstream.searchParams.set('key', env.SHEETS_API_KEY);
  const r = await fetch(upstream, { redirect: 'follow' });
  const body = await r.text();
  if (!r.ok || body === 'forbidden' || body === 'unknown sheet' || body === 'missing sheet') {
    return json({ error: 'upstream', status: r.status }, 502);
  }

  const res = new Response(body, {
    headers: { 'content-type': 'text/csv; charset=utf-8', 'cache-control': `max-age=${CACHE_SECONDS}` },
  });
  ctx.waitUntil(cache.put(cacheKey, res.clone()));
  return withPrivateHeaders(res);
}

// Contador público de la home (sin autenticación): solo entrega el total de pedidos únicos
// (canal + N° de orden, mismas reglas que HB Control) redondeado hacia abajo a la centena.
// Nunca expone filas de ventas. Caché de 6 horas.
async function pedidosPublicos(env, ctx) {
  if (!env.SHEETS_API_URL || !env.SHEETS_API_KEY) return json({ error: 'not_configured' }, 503);
  const cacheKey = new Request('https://cache.internal/public/pedidos');
  const hit = await caches.default.match(cacheKey);
  if (hit) return hit;
  const upstream = new URL(env.SHEETS_API_URL);
  upstream.searchParams.set('sheet', 'ventas');
  upstream.searchParams.set('key', env.SHEETS_API_KEY);
  const r = await fetch(upstream, { redirect: 'follow' });
  const body = await r.text();
  if (!r.ok || body === 'forbidden' || body === 'unknown sheet' || body === 'missing sheet') return json({ error: 'upstream' }, 502);
  const ordenes = new Set();
  for (const line of body.split('\n').slice(1)) {
    if (!line.trim()) continue;
    const c = csvLinea(line);
    if (c.length < 20 || !(c[4] || '').trim() || numero(c[11]) <= 0) continue;
    ordenes.add((c[5] || '').trim() + '|' + (c[6] || ''));
  }
  const pedidos = Math.floor(ordenes.size / 100) * 100;
  const res = new Response(JSON.stringify({ pedidos }), {
    headers: { 'content-type': 'application/json', 'cache-control': 'public, max-age=21600' },
  });
  if (pedidos > 0) ctx.waitUntil(caches.default.put(cacheKey, res.clone()));
  return res;
}

// Precios vigentes (IVA incluido) por SKU, leídos de las tiendas Shopify públicas. Caché de 10 minutos.
const TIENDAS = ['https://memokoala.com', 'https://montanagrills.cl'];
async function preciosWeb(ctx) {
  const cacheKey = new Request('https://cache.internal/public/precios-web-v2');
  const hit = await caches.default.match(cacheKey);
  if (hit) return hit;
  const precios = {};
  const errores = [];
  await Promise.all(TIENDAS.map(async (base) => {
    try {
      let r;
      for (let intento = 0; intento < 3; intento++) {
        r = await fetch(base + '/products.json?limit=250', {
          headers: { accept: 'application/json', 'user-agent': 'Mozilla/5.0 (compatible; HousebrandsCotizador/1.0; +https://housebrands.cl)' },
        });
        if (r.status !== 429) break;
        await new Promise((ok) => setTimeout(ok, 800 * (intento + 1)));
      }
      if (!r.ok) { errores.push(base + ' ' + r.status); return; }
      const d = await r.json();
      for (const p of d.products || []) for (const v of p.variants || []) {
        const sku = String(v.sku || '').trim();
        const precio = Math.round(parseFloat(v.price));
        if (sku && precio > 0) precios[sku] = precio;
      }
    } catch (e) { errores.push(base + ' ' + e.message); }
  }));
  const res = new Response(JSON.stringify({ precios, errores, actualizado: new Date().toISOString() }), {
    headers: { 'content-type': 'application/json', 'cache-control': 'no-store' },
  });
  if (Object.keys(precios).length && !errores.length) ctx.waitUntil(caches.default.put(cacheKey, res.clone()));
  return res;
}

function csvLinea(line) {
  const res = []; let f = '', q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') { if (q && line[i + 1] === '"') { f += '"'; i++; } else q = !q; }
    else if (ch === ',' && !q) { res.push(f.trim()); f = ''; }
    else f += ch;
  }
  res.push(f.trim());
  return res;
}

function numero(v) {
  let s = String(v ?? '').trim().replace(/[$\s%]/g, '');
  if (!s) return 0;
  const dots = (s.match(/\./g) || []).length, commas = (s.match(/,/g) || []).length;
  if (commas >= 1 && dots === 0) s = s.replace(/,/g, '');
  else if (commas === 1 && dots === 1) s = s.replace(',', '');
  else if (commas === 0 && dots > 1) s = s.replace(/\./g, '');
  return parseFloat(s) || 0;
}

// Ventas manuales: el navegador manda {action, venta|orden,canal}; el Worker agrega la clave
// (que nunca llega al navegador) y lo reenvía al doPost de Apps Script.
async function ventaManual(request, env) {
  if (request.method !== 'POST') return json({ error: 'method_not_allowed' }, 405);
  if (!env.ACCESS_TEAM_DOMAIN || !env.ACCESS_AUD || !env.ALLOWED_EMAILS || !env.SHEETS_API_URL || !env.SHEETS_API_KEY) {
    return json({ error: 'not_configured' }, 503);
  }
  const identity = await verifyAccess(request, env).catch(() => null);
  if (!identity) return json({ error: 'unauthorized' }, 401);

  let body;
  try { body = await request.json(); } catch { return json({ ok: false, error: 'json inválido' }, 400); }
  if (!['venta_manual', 'anular_manual', 'enriquecer', 'imp_set', 'sync', 'gastos_set', 'config_set', 'pnl_generar', 'gasto_add', 'gasto_del', 'metas_set', 'pref_set', 'scout_idea_add', 'scout_idea_set', 'amazon_add', 'amazon_clasificar', 'solicitud_add', 'solicitud_set', 'mkt_sync', 'benchmark_add', 'contab_subir', 'contab_borrador'].includes(body.action)) return json({ ok: false, error: 'acción desconocida' }, 400);
  const payload = { ...body, key: env.SHEETS_API_KEY, usuario: identity.email };

  // Apps Script responde 302 a googleusercontent; fetch lo sigue con GET y trae el JSON.
  const r = await fetch(env.SHEETS_API_URL, {
    method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(payload), redirect: 'follow',
  });
  let out;
  try { out = await r.json(); } catch { return json({ ok: false, error: 'respuesta inválida de Sheets', status: r.status }, 502); }
  if (out.ok || body.action === 'sync') {
    const purgar = { imp_set: ['imp'], sync: ['ventas', 'inv', 'estado'], gastos_set: ['gastos'], config_set: ['config'], pnl_generar: ['pnl', 'estado'], gasto_add: ['gdet', 'gastos', 'pnl', 'estado'], gasto_del: ['gdet', 'gastos', 'pnl', 'estado'], metas_set: ['metas'], pref_set: ['prefs'], scout_idea_add: ['ideas'], scout_idea_set: ['ideas'], amazon_add: ['amazon'], amazon_clasificar: ['amazon'], solicitud_add: ['solic'], solicitud_set: ['solic'], benchmark_add: ['bench'], contab_subir: ['contab'], contab_borrador: ['contab'], mkt_sync: ['mktd', 'mktq', 'mktp', 'mktc', 'mkts', 'estado'] }[body.action] || ['ventas'];
    await Promise.all(purgar.map(s => caches.default.delete(new Request('https://cache.internal/erp/' + s))));
  }
  if (out.ok && body.action === 'solicitud_add') out.analisis = await dispararScout(env, out.id);
  return json(out, out.ok ? 200 : 400);
}

// Dispara la tarea programada del Scout para responder un pedido de análisis en el momento.
// Sin token configurado no hace nada: el pedido se responde en la corrida diaria.
async function dispararScout(env, id) {
  if (!env.SCOUT_ROUTINE_ID || !env.SCOUT_FIRE_TOKEN) return { ok: false, motivo: 'sin_configurar' };
  try {
    const r = await fetch(`https://api.anthropic.com/v1/claude_code/routines/${env.SCOUT_ROUTINE_ID}/fire`, {
      method: 'POST',
      headers: { authorization: `Bearer ${env.SCOUT_FIRE_TOKEN}`, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      body: JSON.stringify({ text: `Corrida disparada desde HB Control por el pedido de análisis ${id}. Haz solo el PASO 0, el PASO 2 y AL TERMINAR; no hagas el PASO 1 ni el PASO 1B aunque hoy sea lunes.` }),
    });
    if (r.ok) return { ok: true };
    return { ok: false, motivo: r.status === 429 ? 'limite_por_hora' : 'error_' + r.status };
  } catch (e) {
    return { ok: false, motivo: 'error_red' };
  }
}

function withPrivateHeaders(res) {
  const out = new Response(res.body, res);
  out.headers.set('cache-control', 'private, no-store');
  out.headers.set('x-robots-tag', 'noindex');
  return out;
}

function json(obj, status) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'content-type': 'application/json', 'cache-control': 'no-store' },
  });
}

// ── Validación del JWT de Cloudflare Access (RS256) ─────────────────────────
let certCache = { at: 0, keys: null };

async function verifyAccess(request, env) {
  const token = request.headers.get('cf-access-jwt-assertion');
  if (!token) return null;
  const [h, p, s] = token.split('.');
  if (!h || !p || !s) return null;
  const header = JSON.parse(b64urlText(h));
  const payload = JSON.parse(b64urlText(p));

  const issuer = `https://${env.ACCESS_TEAM_DOMAIN}`;
  const now = Math.floor(Date.now() / 1000);
  const aud = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
  if (payload.iss !== issuer) return null;
  if (!aud.includes(env.ACCESS_AUD)) return null;
  if (!payload.exp || payload.exp < now) return null;

  const allowed = env.ALLOWED_EMAILS.split(',').map((e) => e.trim().toLowerCase()).filter(Boolean);
  const email = String(payload.email || '').toLowerCase();
  if (!allowed.includes(email)) return null;

  const keys = await getKeys(issuer);
  const jwk = keys.find((k) => k.kid === header.kid);
  if (!jwk || header.alg !== 'RS256') return null;
  const key = await crypto.subtle.importKey('jwk', jwk, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
  const ok = await crypto.subtle.verify('RSASSA-PKCS1-v1_5', key, b64urlBytes(s), new TextEncoder().encode(`${h}.${p}`));
  return ok ? { email } : null;
}

async function getKeys(issuer) {
  if (certCache.keys && Date.now() - certCache.at < 10 * 60 * 1000) return certCache.keys;
  const r = await fetch(`${issuer}/cdn-cgi/access/certs`);
  const data = await r.json();
  certCache = { at: Date.now(), keys: data.keys || [] };
  return certCache.keys;
}

function b64urlBytes(str) {
  const b64 = str.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((str.length + 3) % 4);
  const bin = atob(b64);
  return Uint8Array.from(bin, (c) => c.charCodeAt(0));
}

function b64urlText(str) {
  return new TextDecoder().decode(b64urlBytes(str));
}
