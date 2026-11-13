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
const SCOUT_ENC = ['Fecha', 'Grupo', 'Keyword', 'Posición', 'Resultados', 'Precio p25', 'Precio mediana', 'Precio p75', 'Vendedores'];
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
    filas.push([fecha, grupo, kw, pos, m.total, m.p25, m.med, m.p75, m.vendedores]);
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

// Resultados totales, rango de precios (p25, mediana, p75) y vendedores distintos en los primeros 50 resultados.
function _scoutMedir_(kw) {
  const vacio = { total: '', p25: '', med: '', p75: '', vendedores: '' };
  try {
    const j = _mlGet_('/sites/MLC/search?limit=50&q=' + encodeURIComponent(kw));
    const precios = (j.results || []).map(r => Number(r.price)).filter(p => p > 0).sort((a, b) => a - b);
    const q = f => precios.length ? Math.round(precios[Math.min(precios.length - 1, Math.floor(f * precios.length))]) : '';
    const vend = new Set((j.results || []).map(r => r.seller && r.seller.id).filter(Boolean));
    return { total: (j.paging && j.paging.total) || '', p25: q(0.25), med: q(0.5), p75: q(0.75), vendedores: vend.size || '' };
  } catch (e) { return vacio; }
}

// Ejecutar una vez: corre el radar todos los lunes temprano.
function activarScoutSemanal() {
  ScriptApp.getProjectTriggers().filter(t => ['scoutSemanal', 'scoutProgramado'].indexOf(t.getHandlerFunction()) !== -1).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scoutProgramado').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  console.log('Scout semanal activado (lunes 6:00).');
}
