// =============================================================
// HOUSEBRANDS — ScoutAmazon.gs
// Rankings de Amazon (EE.UU., México, España) y pedidos de análisis del Scout.
//
// Los datos de Amazon no se pueden leer desde Apps Script (Amazon bloquea a los servidores). Los carga
// la tarea programada "Scout Amazon y análisis", que corre en el Chrome del dueño: lee los Más vendidos y
// los Nuevos lanzamientos por categoría y los envía con la acción amazon_add. Después clasifica marca,
// genérico y tipo de producto de lo que más sube (acción amazon_clasificar).
//
// Hoja Scout_Amazon (últimas 8 semanas; lo anterior pasa a Scout_Amazon_Hist):
//   Fecha | Mercado | Lista | Categoría | Posición | ASIN | Producto | Precio | Rating | Reseñas |
//   Posición anterior | Cambio | Marca | Genérico | Tipo de producto
//   Cambio = puestos ganados contra la semana anterior (positivo = sube) o "Nuevo".
// Hoja Scout_Solicitudes: ID | Fecha | Tipo | Pedido | Estado | Resultado | Actualizado
//   Estado: Pendiente → En proceso → Listo. La misma tarea programada responde los pendientes.
// =============================================================

const H_AMZ = 'Scout_Amazon';
const H_AMZ_HIST = 'Scout_Amazon_Hist';
const AMZ_ENC = ['Fecha', 'Mercado', 'Lista', 'Categoría', 'Posición', 'ASIN', 'Producto', 'Precio', 'Rating', 'Reseñas', 'Posición anterior', 'Cambio', 'Marca', 'Genérico', 'Tipo de producto'];
const AMZ_SEMANAS = 8;
const H_SOL = 'Scout_Solicitudes';
const SOL_ENC = ['ID', 'Fecha', 'Tipo', 'Pedido', 'Estado', 'Resultado', 'Actualizado'];

function _amzFecha_(v) { return v instanceof Date ? Utilities.formatDate(v, 'America/Santiago', 'yyyy-MM-dd') : String(v).slice(0, 10); }

// filas = [{mercado, lista, categoria, pos, asin, titulo, precio, rating, resenas}], fecha 'yyyy-mm-dd' (opcional)
function amazonAgregar_(filas, fecha) {
  if (!Array.isArray(filas) || !filas.length) throw new Error('sin filas');
  if (filas.length > 3000) throw new Error('demasiadas filas');
  fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(fecha || '')) ? fecha : _smFecha_();
  const h = _smHoja_(H_AMZ, AMZ_ENC);
  const datos = _smDatos_(h, AMZ_ENC.length);
  // Si ya se cargó esta fecha (reintento), se reemplaza
  const resto = datos.filter(r => _amzFecha_(r[0]) !== fecha);
  const fechasPrev = [...new Set(resto.map(r => _amzFecha_(r[0])))].filter(f => f < fecha).sort();
  const fPrev = fechasPrev[fechasPrev.length - 1];
  const prev = {}, clasif = {};
  resto.forEach(r => {
    if (_amzFecha_(r[0]) === fPrev) prev[[r[1], r[2], r[3], r[5]].join('|')] = Number(r[4]) || 0;
    if (r[12] || r[13]) clasif[r[5]] = [r[12], r[13], r[14]];     // la clasificación de un ASIN se hereda
  });
  const nuevas = filas.map(f => {
    const asin = String(f.asin || '').trim().slice(0, 12);
    const pos = Math.round(Number(f.pos) || 0);
    if (!asin || !pos) return null;
    const k = [f.mercado, f.lista, f.categoria, asin].join('|');
    const p = prev[k] || '';
    const c = clasif[asin] || ['', '', ''];
    return [fecha, String(f.mercado || '').slice(0, 10), String(f.lista || '').slice(0, 30), String(f.categoria || '').slice(0, 60), pos, asin,
      String(f.titulo || '').slice(0, 200), String(f.precio || '').slice(0, 30), Number(f.rating) || '', Number(f.resenas) || '',
      p, p ? p - pos : (fPrev ? 'Nuevo' : ''), c[0], c[1], c[2]];
  }).filter(Boolean);
  // Reescribe: semanas recientes + esta carga; lo más antiguo pasa al histórico
  const todas = resto.concat(nuevas);
  const fechas = [...new Set(todas.map(r => _amzFecha_(r[0])))].sort();
  const corte = fechas.length > AMZ_SEMANAS ? fechas[fechas.length - AMZ_SEMANAS] : '';
  const viejas = corte ? todas.filter(r => _amzFecha_(r[0]) < corte) : [];
  const quedan = corte ? todas.filter(r => _amzFecha_(r[0]) >= corte) : todas;
  if (viejas.length) {
    const hh = _smHoja_(H_AMZ_HIST, AMZ_ENC);
    hh.getRange(hh.getLastRow() + 1, 1, viejas.length, AMZ_ENC.length).setValues(viejas);
  }
  if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, AMZ_ENC.length).clearContent();
  if (quedan.length) {
    h.getRange(2, 1, quedan.length, 1).setNumberFormat('@');
    h.getRange(2, 1, quedan.length, AMZ_ENC.length).setValues(quedan.map(r => { r[0] = _amzFecha_(r[0]); return r; }));
  }
  return { ok: true, filas: nuevas.length, comparadaCon: fPrev || null };
}

// items = [{asin, marca, generico ('Sí'|'No'), tipo}] → completa las filas de ese ASIN
function amazonClasificar_(items) {
  if (!Array.isArray(items) || !items.length) throw new Error('sin items');
  const h = _smHoja_(H_AMZ, AMZ_ENC);
  const datos = _smDatos_(h, AMZ_ENC.length);
  const mapa = {};
  items.forEach(i => { if (i && i.asin) mapa[String(i.asin).trim()] = i; });
  let n = 0;
  datos.forEach(r => {
    const i = mapa[r[5]];
    if (!i) return;
    r[12] = String(i.marca || '').slice(0, 60);
    r[13] = i.generico === 'Sí' || i.generico === true ? 'Sí' : 'No';
    r[14] = String(i.tipo || '').slice(0, 80);
    n++;
  });
  if (n) h.getRange(2, 13, datos.length, 3).setValues(datos.map(r => [r[12], r[13], r[14]]));
  return { ok: true, filas: n };
}

function solicitudAgregar_(d) {
  const pedido = String(d.pedido || '').trim();
  if (pedido.length < 3) throw new Error('escribe qué quieres que analice');
  const tipo = ['Producto', 'Nicho', 'Categoría'].indexOf(d.tipo) !== -1 ? d.tipo : 'Producto';
  const h = _smHoja_(H_SOL, SOL_ENC);
  const id = 'S' + Utilities.formatDate(new Date(), 'America/Santiago', 'yyMMddHHmmss');
  h.appendRow([id, _smFecha_(), tipo, pedido.slice(0, 500), 'Pendiente', '', Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH:mm')]);
  return { ok: true, id: id };
}

function solicitudActualizar_(d) {
  const h = _smHoja_(H_SOL, SOL_ENC);
  const datos = _smDatos_(h, 1);
  const i = datos.findIndex(r => String(r[0]) === String(d.id));
  if (i === -1) throw new Error('no se encontró el pedido');
  if (d.estado !== undefined) {
    if (['Pendiente', 'En proceso', 'Listo'].indexOf(d.estado) === -1) throw new Error('estado inválido');
    h.getRange(i + 2, 5).setValue(d.estado);
  }
  if (d.resultado !== undefined) h.getRange(i + 2, 6).setValue(String(d.resultado).slice(0, 40000));
  h.getRange(i + 2, 7).setValue(Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH:mm'));
  return { ok: true };
}
