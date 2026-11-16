// =============================================================
// HOUSEBRANDS — ScoutAmazon.gs
// Rankings de Amazon (EE.UU., México, España) y pedidos de análisis del Scout.
//
// Los datos de Amazon no se pueden leer desde Apps Script (Amazon bloquea a los servidores). Los carga
// la tarea programada "Scout Amazon y análisis", que corre en el Chrome del dueño: HB Control abre una
// pestaña de Amazon, la tarea lee ahí los Más vendidos y los Nuevos lanzamientos por categoría y la
// pestaña se los envía a HB Control (postMessage), que los guarda con la acción amazon_add. Después la
// tarea clasifica marca, genérico y tipo de producto de lo que más sube (acción amazon_clasificar).
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

// Defensa contra lecturas pegadas de Amazon (ej. "3.79434,99 €" = 3.794 reseñas + 34,99 €):
// un precio sin moneda o que no calza con el formato, o más de 2.000.000 de reseñas, se guarda vacío.
// La moneda es obligatoria: un número suelto ("900") es casi siempre un trozo de las reseñas.
const AMZ_PRECIO_OK = /^(?:(?:CLP|US\$|MX\$|\$|€)\s?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?|\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?\s?€)$/;
const AMZ_MAX_RESENAS = 2000000;
function _amzPrecio_(v) { const p = String(v || '').replace(/\s+/g, ' ').trim().slice(0, 30); return AMZ_PRECIO_OK.test(p) ? p : ''; }
function _amzResenas_(v) { const n = Math.round(Number(v) || 0); return n > 0 && n <= AMZ_MAX_RESENAS ? n : ''; }

// filas = [{mercado, lista, categoria, pos, asin, titulo, precio, rating, resenas}], fecha 'yyyy-mm-dd' (opcional)
function amazonAgregar_(filas, fecha) {
  if (!Array.isArray(filas) || !filas.length) throw new Error('sin filas');
  if (filas.length > 3000) throw new Error('demasiadas filas');
  fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(fecha || '')) ? fecha : _smFecha_();
  const h = _smHoja_(H_AMZ, AMZ_ENC);
  const datos = _smDatos_(h, AMZ_ENC.length);
  // Si ya se cargó esta fecha para estos mercados (reintento), se reemplaza; los otros mercados del día se mantienen
  const mercados = new Set(filas.map(f => String(f.mercado || '')));
  const resto = datos.filter(r => !(_amzFecha_(r[0]) === fecha && mercados.has(String(r[1]))));
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
      String(f.titulo || '').slice(0, 200), _amzPrecio_(f.precio), (Number(f.rating) > 0 && Number(f.rating) <= 5) ? Number(f.rating) : '', _amzResenas_(f.resenas),
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
    h.getRange(2, 8, quedan.length, 1).setNumberFormat('@');   // precio como texto: Sheets convertiría "$453.96" en número
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

// Ejecutar una vez desde el editor: vacía precio y reseñas inválidos ya guardados (lecturas pegadas
// del snippet anterior). Por defecto revisa las cargas del 14 y del 16 de noviembre de 2026.
// Si el precio no es confiable (inválido o vacío), las reseñas tampoco lo son: salieron pegadas al precio.
function amazonLimpiarInvalidos(fechas) {
  fechas = Array.isArray(fechas) && fechas.length ? fechas : ['2026-11-14', '2026-11-16'];
  const h = _smHoja_(H_AMZ, AMZ_ENC);
  const datos = _smDatos_(h, AMZ_ENC.length);
  // El precio se lee como se ve en la hoja: si Sheets lo guardó como número, el texto conserva la moneda
  const vistos = datos.length ? h.getRange(2, 8, datos.length, 1).getDisplayValues() : [];
  datos.forEach((r, i) => { r[7] = vistos[i][0]; });
  let precios = 0, resenas = 0, recuperados = 0;
  const salida = datos.map(r => {
    if (fechas.indexOf(_amzFecha_(r[0])) === -1) return [r[7], r[9]];
    let p = String(r[7]).trim(), n = r[9];
    // México: Sheets guardó "$453.96" como número y perdió el signo; ese número sí es el precio
    if (r[1] === 'México' && /^\d+(\.\d+)?$/.test(p)) { p = '$' + Number(p).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); recuperados++; }
    if (p !== '' && !_amzPrecio_(p)) { p = ''; precios++; }
    if (n !== '' && (p === '' || _amzResenas_(n) === '')) { n = ''; resenas++; }
    return [p, n];
  });
  if (datos.length) {
    h.getRange(2, 8, datos.length, 1).setNumberFormat('@').setValues(salida.map(x => [x[0]]));
    h.getRange(2, 10, datos.length, 1).setValues(salida.map(x => [x[1]]));
  }
  const msg = 'Scout_Amazon: ' + precios + ' precios y ' + resenas + ' reseñas inválidos vaciados; ' + recuperados + ' precios de México recuperados con su signo (' + fechas.join(', ') + ')';
  console.log(msg);
  return { ok: true, precios: precios, resenas: resenas, recuperados: recuperados };
}

// =============================================================
// Benchmark de creadores de marcas (Builder, UnionX): una foto mensual del catálogo de cada marca.
// Productos activos por marca = cuánto apostó la empresa por esa marca; los nuevos en 6 meses = si sigue invirtiendo.
// Hoja Scout_Benchmark: Fecha | Empresa | Marca | Propia | Productos | Variantes | Precio mediano | Nuevos 6 meses | Categorías
// La carga la hace la tarea programada (lee /products.json de cada tienda en Chrome) con la acción benchmark_add.
// =============================================================
const H_BENCH = 'Scout_Benchmark';
const BENCH_ENC = ['Fecha', 'Empresa', 'Marca', 'Propia', 'Productos', 'Variantes', 'Precio mediano', 'Nuevos 6 meses', 'Categorías'];

// filas = [{empresa, marca, propia ('Sí'|'No'), productos, variantes, precio, nuevos, categorias}], fecha 'yyyy-mm-dd' (opcional)
function benchmarkAgregar_(filas, fecha) {
  if (!Array.isArray(filas) || !filas.length) throw new Error('sin filas');
  if (filas.length > 500) throw new Error('demasiadas filas');
  fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(fecha || '')) ? fecha : _smFecha_();
  const h = _smHoja_(H_BENCH, BENCH_ENC);
  const empresas = new Set(filas.map(f => String(f.empresa || '').trim()));
  // Reintento del mismo día y empresa: se reemplaza
  const resto = _smDatos_(h, BENCH_ENC.length).filter(r => !(_amzFecha_(r[0]) === fecha && empresas.has(String(r[1]))));
  const nuevas = filas.map(f => {
    const marca = String(f.marca || '').trim().slice(0, 60);
    if (!marca || !f.empresa) return null;
    return [fecha, String(f.empresa).trim().slice(0, 40), marca, f.propia === 'No' ? 'No' : 'Sí',
      Math.round(Number(f.productos) || 0), Math.round(Number(f.variantes) || 0), Math.round(Number(f.precio) || 0) || '',
      Math.round(Number(f.nuevos) || 0), String(f.categorias || '').slice(0, 120)];
  }).filter(Boolean);
  const todas = resto.map(r => { r[0] = _amzFecha_(r[0]); return r; }).concat(nuevas);
  if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, BENCH_ENC.length).clearContent();
  if (todas.length) {
    h.getRange(2, 1, todas.length, 1).setNumberFormat('@');
    h.getRange(2, 1, todas.length, BENCH_ENC.length).setValues(todas);
  }
  return { ok: true, filas: nuevas.length };
}
