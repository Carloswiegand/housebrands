// =============================================================
// HOUSEBRANDS — ScoutMercados.gs
// Radar multipaís y seguimiento de ideas del Scout (costo cero: solo la API de MercadoLibre).
//
// Cada lunes (scoutMercadosProgramado, 6:30):
//  1. Tendencias de búsqueda por rubro en Chile, Argentina, México y Colombia → hoja Scout_Tendencias
//     (grupo "<País> · <Rubro>").
//  2. Brechas: búsquedas que están en tendencia afuera 2 semanas seguidas y no en Chile → hoja
//     Scout_Oportunidades, con la competencia que hoy tienen en el catálogo chileno. Cada brecha nueva se
//     registra sola como idea en seguimiento.
//  3. Seguimiento: cada idea activa se mide en Chile (posición en tendencias, productos y marcas en el
//     catálogo) → hoja Scout_Seguimiento. Veredicto automático:
//       "Se cumplió"    = entró a las tendencias de Chile o su catálogo creció ≥ 30% sobre la base.
//       "No se cumplió" = llegó el plazo sin cumplirse.
//
// Hoja Scout_Ideas: ID | Creada | Keyword | Rubro | Origen | Hipótesis | Marca destino | Plazo |
//   Base productos CL | Base marcas CL | Estado | Veredicto | Fecha veredicto | Notas
// Estado: "En seguimiento" | "Evaluando" | "Descartada" | "Importada". Solo "En seguimiento" y "Evaluando" se miden.
// =============================================================

const SM_SITIOS = { MLC: 'Chile', MLA: 'Argentina', MLM: 'México', MCO: 'Colombia' };
const SM_RUBROS = [
  ['Bebé', 'silla de comer bebe'],
  ['Cocina', 'organizador de cocina'],
  ['Hogar', 'organizador de closet'],
  ['Decoración', 'cuadros decorativos'],
  ['Mascotas', 'comedero para gato'],
  ['Adulto mayor', 'andador adulto mayor'],
  ['Parrilla', 'set parrilla'],
];
const SM_TOP = 20;              // búsquedas por rubro y país
const SM_MAX_BRECHAS = 15;      // brechas nuevas medidas por semana (cada una es una llamada al catálogo)
const SM_PLAZO_DIAS = 84;       // 12 semanas para que la hipótesis se cumpla
const SM_CRECIMIENTO = 1.3;     // +30% de productos en el catálogo chileno = se cumplió

const H_IDEAS = 'Scout_Ideas';
const IDEAS_ENC = ['ID', 'Creada', 'Keyword', 'Rubro', 'Origen', 'Hipótesis', 'Marca destino', 'Plazo', 'Base productos CL', 'Base marcas CL', 'Estado', 'Veredicto', 'Fecha veredicto', 'Notas'];
const H_SEG = 'Scout_Seguimiento';
const SEG_ENC = ['Fecha', 'ID', 'Keyword', 'Posición tendencias CL', 'Productos CL', 'Marcas CL'];
const H_OPOR = 'Scout_Oportunidades';
const OPOR_ENC = ['Fecha', 'Rubro', 'Keyword', 'Dónde está en tendencia', 'Semanas afuera', 'Productos CL', 'Marcas CL', 'Marcas top CL'];

function _smHoja_(nombre, enc) {
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(nombre);
  if (!h) { h = ss.insertSheet(nombre); h.appendRow(enc); h.setFrozenRows(1); h.getRange(1, 1, 1, enc.length).setFontWeight('bold'); }
  return h;
}
function _smDatos_(h, n) { return h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, n).getValues() : []; }
function _smNorm_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim(); }
function _smFecha_(d) { return Utilities.formatDate(d || new Date(), 'America/Santiago', 'yyyy-MM-dd'); }

// Categoría de MercadoLibre con tendencias para un rubro en un país. Las categorías hoja a veces no tienen
// tendencias (404): se sube por el árbol hasta encontrar una que sí. Se guarda para no buscarla cada semana.
function _smCategoria_(sitio, semilla) {
  const props = PropertiesService.getScriptProperties();
  const clave = 'smcat_' + sitio + '_' + _smNorm_(semilla).replace(/[^a-z0-9]/g, '').slice(0, 30);
  const guardada = props.getProperty(clave);
  if (guardada) return JSON.parse(guardada);
  const d = _mlGet_('/sites/' + sitio + '/domain_discovery/search?limit=1&q=' + encodeURIComponent(semilla));
  if (!d || !d[0] || !d[0].category_id) return null;
  const cat = _mlGet_('/categories/' + d[0].category_id);
  const camino = (cat.path_from_root || []).slice().reverse();   // de la más específica a la más general
  for (let i = 0; i < camino.length - 1; i++) {                   // nunca la raíz (demasiado amplia)
    try {
      const t = _mlGet_('/trends/' + sitio + '/' + camino[i].id);
      if (Array.isArray(t) && t.length) {
        const r = { id: camino[i].id, nombre: camino[i].name };
        props.setProperty(clave, JSON.stringify(r));
        return r;
      }
    } catch (e) { /* 404: probar un nivel más arriba */ }
  }
  return null;
}

function scoutMercados() {
  const fecha = _smFecha_();
  const hT = _smHoja_(H_SCOUT, SCOUT_ENC);
  const filasT = [];
  const ahora = {};        // rubro → sitio → [keywords en orden]
  const errores = [];

  // 1. Tendencias por rubro y país
  SM_RUBROS.forEach(([rubro, semilla]) => {
    ahora[rubro] = {};
    Object.keys(SM_SITIOS).forEach(sitio => {
      try {
        const cat = _smCategoria_(sitio, semilla);
        if (!cat) { errores.push(sitio + ' ' + rubro + ': sin categoría'); return; }
        const t = _mlGet_('/trends/' + sitio + '/' + cat.id).filter(x => x && x.keyword).slice(0, SM_TOP);
        ahora[rubro][sitio] = t.map(x => x.keyword);
        t.forEach((x, i) => filasT.push([fecha, SM_SITIOS[sitio] + ' · ' + rubro, x.keyword, i + 1, '', '', '', '', '', '']));
      } catch (e) { errores.push(sitio + ' ' + rubro + ': ' + e.message); }
    });
  });

  // Semana anterior, para exigir 2 semanas seguidas afuera
  const previos = {};
  const datosT = _smDatos_(hT, 4);
  const fechasPrev = [...new Set(datosT.map(r => r[0] instanceof Date ? _smFecha_(r[0]) : String(r[0])))].filter(f => f < fecha).sort();
  const fPrev = fechasPrev[fechasPrev.length - 1];
  datosT.forEach(r => {
    const f = r[0] instanceof Date ? _smFecha_(r[0]) : String(r[0]);
    if (f === fPrev) previos[String(r[1]) + '|' + _smNorm_(r[2])] = Number(r[3]) || 0;
  });
  if (filasT.length) hT.getRange(hT.getLastRow() + 1, 1, filasT.length, SCOUT_ENC.length).setValues(filasT);

  // 2. Brechas: en tendencia afuera dos semanas seguidas, ausentes en Chile
  const hI = _smHoja_(H_IDEAS, IDEAS_ENC);
  const ideas = _smDatos_(hI, IDEAS_ENC.length);
  const yaIdea = new Set(ideas.map(r => _smNorm_(r[2])));
  const hO = _smHoja_(H_OPOR, OPOR_ENC);
  const yaOpor = new Set(_smDatos_(hO, 3).map(r => _smNorm_(r[2])));
  const candidatas = [];
  SM_RUBROS.forEach(([rubro]) => {
    const cl = new Set((ahora[rubro].MLC || []).map(_smNorm_));
    const porKw = {};
    ['MLA', 'MLM', 'MCO'].forEach(sitio => {
      (ahora[rubro][sitio] || []).forEach((kw, i) => {
        const n = _smNorm_(kw);
        if (cl.has(n)) return;
        const prev = previos[SM_SITIOS[sitio] + ' · ' + rubro + '|' + n];
        if (!prev) return;                                   // solo una semana: todavía no
        porKw[n] = porKw[n] || { kw: kw, rubro: rubro, donde: [], mejor: 99 };
        porKw[n].donde.push(SM_SITIOS[sitio] + ' #' + (i + 1));
        porKw[n].mejor = Math.min(porKw[n].mejor, i + 1);
      });
    });
    Object.keys(porKw).forEach(n => { if (!yaOpor.has(n) && !yaIdea.has(n)) candidatas.push(porKw[n]); });
  });
  // Primero las que están en más países y más arriba
  candidatas.sort((a, b) => (b.donde.length - a.donde.length) || (a.mejor - b.mejor));
  const nuevasO = [], nuevasI = [];
  candidatas.slice(0, SM_MAX_BRECHAS).forEach(c => {
    const m = _scoutMedir_(c.kw);
    nuevasO.push([fecha, c.rubro, c.kw, c.donde.join(', '), 2, m.total, m.vendedores, m.top]);
    nuevasI.push(_smNuevaIdea_({ keyword: c.kw, rubro: c.rubro, origen: 'Auto · tendencia en ' + c.donde.join(', '), base: m }));
  });
  if (nuevasO.length) hO.getRange(hO.getLastRow() + 1, 1, nuevasO.length, OPOR_ENC.length).setValues(nuevasO);
  if (nuevasI.length) hI.getRange(hI.getLastRow() + 1, 1, nuevasI.length, IDEAS_ENC.length).setValues(nuevasI);

  // 3. Seguimiento de las ideas activas
  scoutSeguimiento_(ahora);

  if (errores.length) console.log('Scout mercados, avisos: ' + errores.join(' | '));
  return filasT.length;
}

function _smNuevaIdea_(o) {
  const id = 'I' + Utilities.formatDate(new Date(), 'America/Santiago', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  const plazo = new Date(Date.now() + (Number(o.plazoDias) || SM_PLAZO_DIAS) * 864e5);
  const m = o.base || {};
  return [id, _smFecha_(), String(o.keyword).slice(0, 80), String(o.rubro || '').slice(0, 40), String(o.origen || 'Manual').slice(0, 120),
    String(o.hipotesis || 'Entra a las tendencias de Chile o su catálogo chileno crece 30% en 12 semanas').slice(0, 300),
    String(o.marca || '').slice(0, 40), _smFecha_(plazo), m.total || 0, m.vendedores || 0, o.estado || 'En seguimiento', '', '', String(o.notas || '').slice(0, 300)];
}

// Mide en Chile cada idea activa y pone el veredicto cuando corresponde.
function scoutSeguimiento_(ahora) {
  const hI = _smHoja_(H_IDEAS, IDEAS_ENC);
  const ideas = _smDatos_(hI, IDEAS_ENC.length);
  if (!ideas.length) return 0;
  // Posiciones en las tendencias de Chile esta semana: general + todos los rubros
  const posCL = {};
  try { _mlGet_('/trends/MLC').forEach((x, i) => { const n = _smNorm_(x.keyword); if (!posCL[n]) posCL[n] = i + 1; }); } catch (e) {}
  Object.keys(ahora || {}).forEach(r => (ahora[r].MLC || []).forEach((kw, i) => { const n = _smNorm_(kw); if (!posCL[n] || posCL[n] > i + 1) posCL[n] = i + 1; }));
  const hS = _smHoja_(H_SEG, SEG_ENC);
  const fecha = _smFecha_(), hoy = new Date(), filas = [];
  ideas.forEach((r, i) => {
    if (['En seguimiento', 'Evaluando'].indexOf(String(r[10])) === -1) return;
    const n = _smNorm_(r[2]);
    // Posición: exacta, o la tendencia chilena que contiene la idea (p. ej. "comedero gato automatico")
    let pos = posCL[n] || '';
    if (!pos) Object.keys(posCL).forEach(k => { if (k.indexOf(n) !== -1 && (!pos || posCL[k] < pos)) pos = posCL[k]; });
    const m = _scoutMedir_(r[2]);
    filas.push([fecha, r[0], r[2], pos, m.total || '', m.vendedores || '']);
    if (!r[11]) {
      const base = Number(r[8]) || 0;
      const plazo = r[7] instanceof Date ? r[7] : new Date(String(r[7]) + 'T12:00:00');
      let v = '';
      if (pos) v = 'Se cumplió';
      else if (base > 0 && Number(m.total) >= base * SM_CRECIMIENTO) v = 'Se cumplió';
      else if (hoy >= plazo) v = 'No se cumplió';
      if (v) hI.getRange(i + 2, 12, 1, 2).setValues([[v, fecha]]);
    }
  });
  if (filas.length) hS.getRange(hS.getLastRow() + 1, 1, filas.length, SEG_ENC.length).setValues(filas);
  return filas.length;
}

// ── Acciones desde HB Control (Api.gs) ───────────────────────
function scoutIdeaAgregar_(d) {
  const kw = String(d.keyword || '').trim();
  if (kw.length < 3) throw new Error('escribe la búsqueda o producto a seguir');
  const hI = _smHoja_(H_IDEAS, IDEAS_ENC);
  if (_smDatos_(hI, 3).some(r => _smNorm_(r[2]) === _smNorm_(kw))) throw new Error('esa idea ya está en seguimiento');
  const fila = _smNuevaIdea_({ keyword: kw, rubro: d.rubro, origen: d.origen || 'Manual', hipotesis: d.hipotesis, marca: d.marca, plazoDias: d.plazoDias, notas: d.notas, base: _scoutMedir_(kw) });
  hI.appendRow(fila);
  return { ok: true, id: fila[0] };
}

function scoutIdeaActualizar_(d) {
  const hI = _smHoja_(H_IDEAS, IDEAS_ENC);
  const ideas = _smDatos_(hI, 1);
  const i = ideas.findIndex(r => String(r[0]) === String(d.id));
  if (i === -1) throw new Error('no se encontró la idea');
  const fila = i + 2;
  if (d.estado !== undefined) {
    if (['En seguimiento', 'Evaluando', 'Descartada', 'Importada'].indexOf(d.estado) === -1) throw new Error('estado inválido');
    hI.getRange(fila, 11).setValue(d.estado);
  }
  if (d.veredicto !== undefined) {
    if (['', 'Se cumplió', 'No se cumplió'].indexOf(d.veredicto) === -1) throw new Error('veredicto inválido');
    hI.getRange(fila, 12, 1, 2).setValues([[d.veredicto, d.veredicto ? _smFecha_() : '']]);
  }
  if (d.marca !== undefined) hI.getRange(fila, 7).setValue(String(d.marca).slice(0, 40));
  if (d.notas !== undefined) hI.getRange(fila, 14).setValue(String(d.notas).slice(0, 300));
  return { ok: true };
}

function scoutMercadosProgramado() { correr_('Scout mercados', scoutMercados); }

// Ejecutar una vez: deja el radar multipaís corriendo los lunes a las 6:30 (después del radar de Chile).
function activarScoutMercados() {
  ScriptApp.getProjectTriggers().filter(t => ['scoutMercados', 'scoutMercadosProgramado'].indexOf(t.getHandlerFunction()) !== -1).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scoutMercadosProgramado').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).nearMinute(30).create();
  console.log('Scout mercados activado (lunes 6:30).');
}
