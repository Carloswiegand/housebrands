// =============================================================
// HOUSEBRANDS — Scout.gs
// Radar semanal de MercadoLibre Chile para el módulo Scout de HB Control.
// Guarda cada semana en la hoja Scout_Tendencias:
//   Fecha | Grupo | Keyword | Posición | Resultados | Precio p25 | Precio mediana | Precio p75 | Vendedores
// - "General": búsquedas más populares de ML Chile (/trends/MLC).
// - Una fila por categoría de las marcas (se detecta sola con domain_discovery).
// - "Mis marcas": las búsquedas propias de MemoKoala y Montana, para medir competencia y precios.
// Con las semanas acumuladas el ERP calcula qué sube, qué baja y qué es nuevo.
// =============================================================

const H_SCOUT = 'Scout_Tendencias';
const SCOUT_ENC = ['Fecha', 'Grupo', 'Keyword', 'Posición', 'Resultados', 'Precio p25', 'Precio mediana', 'Precio p75', 'Vendedores', 'Marcas top'];
const SCOUT_MIS_MARCAS = [
  ['MemoKoala', 'silla de comer portatil bebe'],
  ['MemoKoala', 'mecedora bebe'],
  ['MemoKoala', 'organizador coche bebe'],
  ['Montana', 'cuchillo japones'],
  ['Montana', 'cuchillo chef'],
  ['Montana', 'cuchillo santoku'],
  ['Montana', 'afilador de cuchillos'],
];
const SCOUT_TOP_GENERAL = 30;     // búsquedas generales a medir
const SCOUT_TOP_CATEGORIA = 15;   // búsquedas por categoría

function scoutSemanal() {
  const fecha = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd');
  const filas = [];
  const medir = (grupo, kw, pos) => {
    const m = _scoutMedir_(kw);
    filas.push([fecha, grupo, kw, pos, m.total, m.p25, m.med, m.p75, m.vendedores, m.top]);
  };

  // 1. Tendencias generales de ML Chile
  const gen = _scoutTrends_('');
  gen.slice(0, SCOUT_TOP_GENERAL).forEach((t, i) => medir('General', t.keyword, i + 1));

  // 2. Tendencias por las categorías donde venden las marcas
  const cats = {};
  SCOUT_MIS_MARCAS.forEach(([, kw]) => {
    try {
      const d = _mlGet_('/sites/MLC/domain_discovery/search?limit=1&q=' + encodeURIComponent(kw));
      if (d && d[0] && d[0].category_id) cats[d[0].category_id] = d[0].category_name || d[0].domain_name || d[0].category_id;
    } catch (e) { /* sigue con las demás */ }
  });
  Object.keys(cats).forEach(cid => {
    _scoutTrends_(cid).slice(0, SCOUT_TOP_CATEGORIA).forEach((t, i) => medir('Categoría: ' + cats[cid], t.keyword, i + 1));
  });

  // 3. Búsquedas propias de cada marca
  SCOUT_MIS_MARCAS.forEach(([marca, kw]) => medir('Mis marcas: ' + marca, kw, ''));

  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(H_SCOUT);
  if (!h) { h = ss.insertSheet(H_SCOUT); h.appendRow(SCOUT_ENC); h.setFrozenRows(1); }
  h.getRange(1, 1, 1, SCOUT_ENC.length).setValues([SCOUT_ENC]);   // agrega columnas nuevas al encabezado
  if (filas.length) h.getRange(h.getLastRow() + 1, 1, filas.length, SCOUT_ENC.length).setValues(filas);
  return filas.length;
}

function scoutProgramado() { correr_('Scout semanal', scoutSemanal); }

function _scoutTrends_(categoria) {
  try {
    const j = _mlGet_('/trends/MLC' + (categoria ? '/' + categoria : ''));
    return Array.isArray(j) ? j.filter(t => t && t.keyword) : [];
  } catch (e) { return []; }
}

// Competencia medida en el catálogo de MercadoLibre (/products/search). La búsqueda pública
// (/sites/MLC/search) y los rankings de más vendidos (/highlights) responden 403 para esta app desde 2025,
// así que no hay precios por API: se registran vacíos y se revisan a mano.
//   Resultados = productos de catálogo activos para la búsqueda
//   Vendedores = marcas distintas entre los primeros 50 productos (columna reutilizada)
//   Marcas top = las 5 marcas que más se repiten en esos 50
function _scoutMedir_(kw) {
  const vacio = { total: '', p25: '', med: '', p75: '', vendedores: '', top: '' };
  try {
    const j = _mlGet_('/products/search?status=active&site_id=MLC&limit=50&q=' + encodeURIComponent(kw));
    const cuenta = {};
    (j.results || []).forEach(r => {
      const b = ((r.attributes || []).find(a => a.id === 'BRAND') || {}).value_name;
      if (b && !/gen[eé]ric|sin marca/i.test(b)) cuenta[b] = (cuenta[b] || 0) + 1;
    });
    const top = Object.keys(cuenta).sort((a, b) => cuenta[b] - cuenta[a]).slice(0, 5).join(' · ');
    return { total: (j.paging && j.paging.total) || '', p25: '', med: '', p75: '', vendedores: Object.keys(cuenta).length || '', top: top };
  } catch (e) { return vacio; }
}

// Ejecutar una vez: corre el radar todos los lunes temprano.
function activarScoutSemanal() {
  ScriptApp.getProjectTriggers().filter(t => ['scoutSemanal', 'scoutProgramado'].indexOf(t.getHandlerFunction()) !== -1).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scoutProgramado').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  console.log('Scout semanal activado (lunes 6:00).');
}
