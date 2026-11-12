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

const SHEETS = new Set(['ventas', 'pnl', 'conv', 'inv']);
const CACHE_SECONDS = 60;

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
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
  if (!['venta_manual', 'anular_manual'].includes(body.action)) return json({ ok: false, error: 'acción desconocida' }, 400);
  const payload = { ...body, key: env.SHEETS_API_KEY, usuario: identity.email };

  // Apps Script responde 302 a googleusercontent; fetch lo sigue con GET y trae el JSON.
  const r = await fetch(env.SHEETS_API_URL, {
    method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(payload), redirect: 'follow',
  });
  let out;
  try { out = await r.json(); } catch { return json({ ok: false, error: 'respuesta inválida de Sheets', status: r.status }, 502); }
  if (out.ok) await caches.default.delete(new Request('https://cache.internal/erp/ventas'));
  return json(out, out.ok ? 200 : 400);
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
