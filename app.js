/* ==========================================================================
 * Revisión de topes · Horario Ingeniería
 *
 * Sitio estático (GitHub Pages). Lee:
 *   - data/config.json  → periodo, URL del Web App (descarga Excel y consulta RUT)
 *   - data/data.json    → filas de la hoja HORARIO (publicadas desde Apps Script)
 *
 * La relación RUT alumno → NRC inscritos NO se publica en el repositorio:
 * se consulta un RUT a la vez al Web App de Apps Script (action=nrcsPorRut).
 * ========================================================================== */

const FILTER_CONFIG = [
  { id: 'area', label: 'ÁREA' },
  { id: 'materia', label: 'MATERIA' },
  { id: 'curso', label: 'CURSO' },
  { id: 'seccion', label: 'SECCIÓN' },
  { id: 'titulo', label: 'TÍTULO' },
  { id: 'nrc', label: 'NRC' },
  { id: 'conector', label: 'CONECTOR DE LIGA' },
  { id: 'lc', label: 'LC' },
  { id: 'sala', label: 'SALA' },
  { id: 'tipo', label: 'TIPO DE REUNIÓN' },
  { id: 'profesor', label: 'PROFESOR' }
];

const DAYS = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes'];
const DAY_LABEL = { lunes: 'Lunes', martes: 'Martes', miercoles: 'Miércoles', jueves: 'Jueves', viernes: 'Viernes' };
const DAY_COLUMN_INDEX = { lunes: 9, martes: 10, miercoles: 11, jueves: 12, viernes: 13 };
const INICIO_COLUMN_INDEX = 14;
const RUT_LOOKUP_TIMEOUT_MS = 30000;

let STATIC_CONFIG = null;
let STATIC_DATA = null;
let OWN_WEEKLY_BLOCKS = {};        // NRC → bloques semanales (de todo el horario)
let DATA_NRC_SET = new Set();      // NRC presentes en data.json (normalizados)

let BLOQUEOS = null;               // data/bloqueos.json → { cursos: { CODIGO: {CLAS, AYUD, LAB, numEval} } }
let LC_BLOQUEO_CODE = {};          // LC/conector → código con información de bloqueo
let conflictReviewActive = false;
let reviewMode = 'basic';          // 'basic' | 'bloqueos'
let currentRenderedRows = [];
let currentRowElements = [];
let rawQueryRows = [];
let availableOptionsState = {};
let selectedValuesState = {};
let dayNonEmptyFilters = { lunes: false, martes: false, miercoles: false, jueves: false, viernes: false };
let legendCollapsed = false;
let studentContext = null;         // { rut, inscripciones, conHorario, sinHorario }

/* --------------------------------------------------------------------------
 * Utilidades
 * ------------------------------------------------------------------------ */

function normalizeText(value) {
  return String(value == null ? '' : value)
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .trim()
    .toUpperCase();
}

function normalizeNrc(value) {
  const s = String(value == null ? '' : value).trim();
  if (/^\d+(\.0+)?$/.test(s)) return String(Number(s));
  return s.toUpperCase();
}

function normalizeRut(value) {
  return String(value == null ? '' : value)
    .toUpperCase()
    .replace(/[^0-9K]/g, '');
}

function isValidRut(rutNorm) {
  const m = /^(\d{7,8})([0-9K])$/.exec(rutNorm);
  if (!m) return false;
  const body = m[1];
  let sum = 0;
  let mul = 2;
  for (let i = body.length - 1; i >= 0; i--) {
    sum += Number(body[i]) * mul;
    mul = mul === 7 ? 2 : mul + 1;
  }
  const res = 11 - (sum % 11);
  const dv = res === 11 ? '0' : res === 10 ? 'K' : String(res);
  return dv === m[2];
}

function formatRut(rutNorm) {
  const body = rutNorm.slice(0, -1);
  const dv = rutNorm.slice(-1);
  return body.replace(/\B(?=(\d{3})+(?!\d))/g, '.') + '-' + dv;
}

function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function uniqueSorted(arr) {
  return [...new Set(arr.filter(v => String(v || '').trim() !== ''))]
    .sort((a, b) => String(a).localeCompare(String(b), 'es', { numeric: true }));
}

function parseDateDMY(text) {
  const m = String(text || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])).getTime();
}

function parseTimeRange(text) {
  const m = String(text || '').match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return {
    start: Number(m[1]) * 60 + Number(m[2]),
    end: Number(m[3]) * 60 + Number(m[4]),
    text: `${m[1].padStart(2, '0')}:${m[2]}-${m[3].padStart(2, '0')}:${m[4]}`
  };
}

function timesOverlap(a, b) {
  return Math.max(a.start, b.start) < Math.min(a.end, b.end);
}

/* --------------------------------------------------------------------------
 * Carga de datos
 * ------------------------------------------------------------------------ */

async function loadStaticFiles() {
  const [configRes, dataRes] = await Promise.all([
    fetch('./data/config.json', { cache: 'no-store' }),
    fetch('./data/data.json', { cache: 'no-store' })
  ]);
  if (!configRes.ok) throw new Error('No se pudo cargar config.json');
  if (!dataRes.ok) throw new Error('No se pudo cargar data.json');
  STATIC_CONFIG = await configRes.json();
  STATIC_DATA = await dataRes.json();

  const rows = STATIC_DATA?.rows || [];
  DATA_NRC_SET = new Set(rows.map(r => normalizeNrc(r.meta?.nrc)).filter(Boolean));

  // Bloques semanales propios de cada NRC, calculados sobre TODO el horario
  // (no solo la selección). Sirven para saber si una PRBA está dentro del
  // horario regular de su curso o fuera de él (tarde, otro día, etc.).
  OWN_WEEKLY_BLOCKS = {};
  rows.forEach(row => {
    const rec = buildRecord(row, -1);
    if (!rec || !rec.weekly) return;
    (OWN_WEEKLY_BLOCKS[rec.nrc] = OWN_WEEKLY_BLOCKS[rec.nrc] || []).push(rec);
  });

  // Bloqueos académicos (opcional: si no existe el archivo, el botón avisa).
  BLOQUEOS = null;
  LC_BLOQUEO_CODE = {};
  try {
    const res = await fetch('./data/bloqueos.json', { cache: 'no-store' });
    if (res.ok) BLOQUEOS = await res.json();
  } catch (e) { /* sin bloqueos */ }

  if (BLOQUEOS?.cursos) {
    // Cursos con código antiguo o sin fila en BLOQUEOS heredan los bloqueos de
    // otro curso de su misma LC (o conector de liga) que sí tenga información.
    rows.forEach(row => {
      const code = courseCode(row.meta);
      if (!BLOQUEOS.cursos[code]) return;
      [row.meta.lc, row.meta.conector].forEach(k => {
        const key = normalizeText(k);
        if (key && !LC_BLOQUEO_CODE[key]) LC_BLOQUEO_CODE[key] = code;
      });
    });
  }
}

function courseCode(meta) {
  const materia = normalizeText(meta.materia).replace(/\s+/g, '');
  let curso = normalizeText(meta.curso).replace(/^'/, '').replace(/\s+/g, '');
  if (/^\d+$/.test(curso)) curso = curso.padStart(4, '0');
  return materia + curso;
}

function renderPeriodMeta() {
  const el = document.getElementById('periodMeta');
  const parts = [];
  if (STATIC_CONFIG?.periodo) parts.push(`Periodo ${STATIC_CONFIG.periodo}`);
  if (STATIC_DATA?.generatedAt) {
    const d = new Date(STATIC_DATA.generatedAt);
    if (!isNaN(d)) parts.push(`Actualizado: ${d.toLocaleString('es-CL', { dateStyle: 'short', timeStyle: 'short' })}`);
  }
  el.textContent = parts.join(' · ');
}

/* --------------------------------------------------------------------------
 * Filtros (multiselección con facetas)
 * ------------------------------------------------------------------------ */

function matchExactMulti(rowValue, selectedValues) {
  if (!selectedValues || !selectedValues.length) return true;
  const target = normalizeText(rowValue);
  return selectedValues.some(v => normalizeText(v) === target);
}

function rowMatchesMeta(meta, state, excludeKey = null) {
  for (const cfg of FILTER_CONFIG) {
    if (cfg.id === excludeKey) continue;
    if (!matchExactMulti(meta[cfg.id], state[cfg.id])) return false;
  }
  return true;
}

function buildFacetOptions(allRows, state) {
  const options = {};
  FILTER_CONFIG.forEach(cfg => {
    const compatible = allRows.filter(row => rowMatchesMeta(row.meta, state, cfg.id));
    options[cfg.id] = uniqueSorted(compatible.map(r => r.meta[cfg.id] || ''));
  });
  return options;
}

function getState() {
  const state = {};
  FILTER_CONFIG.forEach(cfg => { state[cfg.id] = [...(selectedValuesState[cfg.id] || [])]; });
  return state;
}

function buildFilters() {
  const grid = document.getElementById('filtersGrid');
  grid.innerHTML = '';
  FILTER_CONFIG.forEach(cfg => {
    selectedValuesState[cfg.id] = selectedValuesState[cfg.id] || [];
    const wrapper = document.createElement('div');
    wrapper.className = 'multiselect';
    wrapper.innerHTML = `
      <label>${cfg.label}</label>
      <button type="button" id="btn-${cfg.id}" class="multi-btn" onclick="togglePanel('${cfg.id}')">Todos</button>
      <div id="panel-${cfg.id}" class="multi-panel">
        <div class="panel-topbar">
          <input id="search-${cfg.id}" class="panel-search" type="text" placeholder="Buscar" oninput="filterPanelOptions('${cfg.id}')" />
          <button type="button" class="panel-clear-btn" onclick="clearSingleFilter('${cfg.id}', event)">Limpiar</button>
        </div>
        <div id="options-${cfg.id}" class="multi-options"></div>
      </div>`;
    grid.appendChild(wrapper);
  });
}

function setFilterOptions(id, values) {
  availableOptionsState[id] = values || [];
  const box = document.getElementById(`options-${id}`);
  const selected = new Set(selectedValuesState[id] || []);
  box.innerHTML = '';
  availableOptionsState[id].forEach(v => {
    const row = document.createElement('label');
    row.className = 'multi-option';
    row.setAttribute('data-label', String(v).toLowerCase());
    const chk = document.createElement('input');
    chk.type = 'checkbox';
    chk.value = v;
    chk.checked = selected.has(v);
    chk.addEventListener('change', () => onFilterChange(id));
    const span = document.createElement('span');
    span.textContent = v;
    row.append(chk, ' ', span);
    box.appendChild(row);
  });
  filterPanelOptions(id);
  updateFilterButton(id);
}

function filterPanelOptions(id) {
  const query = (document.getElementById(`search-${id}`).value || '').toLowerCase().trim();
  document.querySelectorAll(`#options-${id} .multi-option`).forEach(row => {
    const label = row.getAttribute('data-label') || '';
    row.style.display = !query || label.includes(query) ? 'flex' : 'none';
  });
}

function syncSelectionFromDom(id) {
  const available = availableOptionsState[id] || [];
  const domSelected = [...document.querySelectorAll(`#options-${id} input[type="checkbox"]:checked`)].map(el => el.value);
  const hiddenSelected = (selectedValuesState[id] || []).filter(v => !available.includes(v));
  selectedValuesState[id] = [...hiddenSelected, ...domSelected];
}

function updateFilterButton(id) {
  const btn = document.getElementById(`btn-${id}`);
  const selected = selectedValuesState[id] || [];
  if (!selected.length) {
    btn.textContent = 'Todos';
    btn.classList.remove('active');
  } else if (selected.length === 1) {
    btn.textContent = selected[0];
    btn.classList.add('active');
  } else {
    btn.textContent = `${selected.length} seleccionados`;
    btn.classList.add('active');
  }
}

function onFilterChange(id) {
  syncSelectionFromDom(id);
  updateFilterButton(id);
  resetConflictReviewState();
  queryAndRender();
}

function clearSingleFilter(id, event) {
  if (event) { event.preventDefault(); event.stopPropagation(); }
  selectedValuesState[id] = [];
  document.getElementById(`search-${id}`).value = '';
  resetConflictReviewState();
  queryAndRender();
}

function togglePanel(id) {
  document.querySelectorAll('.multi-panel').forEach(p => {
    if (p.id !== `panel-${id}`) p.classList.remove('open');
  });
  document.getElementById(`panel-${id}`).classList.toggle('open');
}

document.addEventListener('click', e => {
  if (!e.target.closest('.multiselect')) {
    document.querySelectorAll('.multi-panel').forEach(p => p.classList.remove('open'));
  }
});

function resetAllFilters() {
  FILTER_CONFIG.forEach(cfg => {
    selectedValuesState[cfg.id] = [];
    const search = document.getElementById(`search-${cfg.id}`);
    if (search) search.value = '';
  });
  dayNonEmptyFilters = { lunes: false, martes: false, miercoles: false, jueves: false, viernes: false };
  updateDayFilterButtons();
}

function limpiar() {
  resetAllFilters();
  resetConflictReviewState();
  queryAndRender();
}

/* --------------------------------------------------------------------------
 * Filtros "Solo <día> con valor"
 * ------------------------------------------------------------------------ */

function applyDayNonEmptyFilters(rows) {
  return (rows || []).filter(row =>
    DAYS.every(day => !dayNonEmptyFilters[day] || String(row.meta?.[day] || '').trim())
  );
}

function updateDayFilterButtons() {
  const mapping = { lunes: 'btnLunes', martes: 'btnMartes', miercoles: 'btnMiercoles', jueves: 'btnJueves', viernes: 'btnViernes' };
  Object.entries(mapping).forEach(([day, id]) => {
    document.getElementById(id)?.classList.toggle('active', !!dayNonEmptyFilters[day]);
  });
}

function toggleDayNonEmpty(day) {
  dayNonEmptyFilters[day] = !dayNonEmptyFilters[day];
  updateDayFilterButtons();
  resetConflictReviewState();
  renderResults(applyDayNonEmptyFilters(rawQueryRows));
}

function clearDayNonEmptyFilters() {
  dayNonEmptyFilters = { lunes: false, martes: false, miercoles: false, jueves: false, viernes: false };
  updateDayFilterButtons();
  resetConflictReviewState();
  renderResults(applyDayNonEmptyFilters(rawQueryRows));
}

/* --------------------------------------------------------------------------
 * Render tabla principal
 * ------------------------------------------------------------------------ */

function renderLegend(rows) {
  const tbody = document.getElementById('legendBody');
  tbody.innerHTML = '';
  if (!rows.length) {
    tbody.innerHTML = '<tr><td>No hay leyenda cargada.</td></tr>';
    return;
  }
  rows.forEach(row => {
    const tr = document.createElement('tr');
    row.cells.forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell.value || '';
      td.style.background = cell.bg || '#ffffff';
      td.style.color = cell.color || '#111827';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

function toggleLegend() {
  legendCollapsed = !legendCollapsed;
  document.getElementById('legendWrap').classList.toggle('hidden', legendCollapsed);
  document.getElementById('legendToggleBtn').textContent = legendCollapsed ? 'Mostrar leyenda' : 'Ocultar leyenda';
}

function renderResults(rows) {
  const tbody = document.getElementById('resultsBody');
  tbody.innerHTML = '';
  currentRenderedRows = rows || [];
  currentRowElements = [];

  if (!currentRenderedRows.length) {
    tbody.innerHTML = '<tr><td colspan="19">No se encontraron resultados.</td></tr>';
  } else {
    const frag = document.createDocumentFragment();
    currentRenderedRows.forEach((row, rowIndex) => {
      const tr = document.createElement('tr');
      tr.dataset.rowIndex = rowIndex;
      row.cells.forEach(cell => {
        const td = document.createElement('td');
        td.textContent = cell.value || '';
        td.style.background = cell.bg || '#ffffff';
        td.style.color = cell.color || '#111827';
        td.dataset.baseBg = cell.bg || '#ffffff';
        tr.appendChild(td);
      });
      frag.appendChild(tr);
      currentRowElements.push(tr);
    });
    tbody.appendChild(frag);
  }

  document.getElementById('meta').textContent = `Resultados encontrados: ${currentRenderedRows.length}`;
  syncScrollWidth();
}

function setLoading(show) {
  document.getElementById('loading').style.display = show ? 'block' : 'none';
}

function queryAndRender() {
  const state = getState();
  const allRows = STATIC_DATA?.rows || [];
  const options = buildFacetOptions(allRows, state);
  FILTER_CONFIG.forEach(cfg => setFilterOptions(cfg.id, options[cfg.id] || []));
  rawQueryRows = allRows.filter(row => rowMatchesMeta(row.meta, state));
  renderResults(applyDayNonEmptyFilters(rawQueryRows));
}

/* --------------------------------------------------------------------------
 * Motor de topes
 *
 * Reglas (para cada par de filas de la selección actual):
 *  - Se ignoran pares del mismo NRC, mismo curso (MATERIA+CURSO, distintas
 *    secciones), misma LC y cursos homologables (misma materia y título,
 *    distinto número de curso).
 *  - Semanal vs semanal (CLAS/AYUD/LAB): mismo día, horas que se cruzan y
 *    periodos INICIO–FIN que se cruzan.
 *  - Evaluación vs evaluación (PRBA/EXAM): misma fecha y horas que se cruzan.
 *  - Evaluación vs semanal: la evaluación cae dentro del periodo de la clase
 *    del otro curso, mismo día y hora, y la evaluación está FUERA del horario
 *    regular de su propio curso (si estuviera dentro, el tope ya aparece como
 *    semanal vs semanal).
 * ------------------------------------------------------------------------ */

function normalizeTipo(value) {
  const t = normalizeText(value);
  if (t === 'LABT' || t === 'LAB' || t === 'TALLER') return 'LAB/TALLER';
  if (t === 'OLIN') return 'CLAS';   // clase online
  if (t === 'AYON') return 'AYUD';   // ayudantía online
  return t;
}

function isWeeklyTipo(t) { return t === 'CLAS' || t === 'AYUD' || t === 'LAB/TALLER'; }
function isEvalTipo(t) { return t === 'PRBA' || t === 'EXAM'; }

function buildRecord(row, index) {
  const m = row.meta || {};
  const tipo = normalizeTipo(m.tipo);
  const slots = [];
  DAYS.forEach(day => {
    const range = parseTimeRange(m[day]);
    if (range) slots.push({ day, range });
  });
  if (!slots.length) return null;
  return {
    index,
    meta: m,
    nrc: normalizeNrc(m.nrc),
    tipo,
    weekly: isWeeklyTipo(tipo),
    eval: isEvalTipo(tipo),
    slots,
    start: parseDateDMY(m.inicio),
    end: parseDateDMY(m.fin)
  };
}

function shouldSkipPair(a, b) {
  const ma = a.meta, mb = b.meta;
  if (a.nrc && a.nrc === b.nrc) return true;

  const matA = normalizeText(ma.materia), matB = normalizeText(mb.materia);
  const curA = normalizeText(ma.curso), curB = normalizeText(mb.curso);
  if (matA && matA === matB && curA && curA === curB) return true;

  const lcA = normalizeText(ma.lc), lcB = normalizeText(mb.lc);
  if (lcA && lcA === lcB) return true;

  const titA = normalizeText(ma.titulo), titB = normalizeText(mb.titulo);
  if (matA && matA === matB && titA && titA === titB) return true;

  return false;
}

function periodsOverlap(a, b) {
  // Si falta alguna fecha, no se descarta el tope (comportamiento conservador).
  if (a.start == null || a.end == null || b.start == null || b.end == null) return true;
  return a.start <= b.end && b.start <= a.end;
}

function dateInside(date, rec) {
  if (date == null || rec.start == null || rec.end == null) return true;
  return date >= rec.start && date <= rec.end;
}

function evalInsideOwnSchedule(ev, slot) {
  const own = OWN_WEEKLY_BLOCKS[ev.nrc] || [];
  return own.some(w =>
    w.slots.some(ws => ws.day === slot.day && timesOverlap(ws.range, slot.range)) &&
    dateInside(ev.start, w)
  );
}

function detectConflicts(rows) {
  const records = rows.map((row, i) => buildRecord(row, i)).filter(Boolean);

  // Agrupar por día: solo se comparan filas que comparten día de la semana.
  const byDay = {};
  records.forEach(rec => {
    rec.slots.forEach(slot => {
      (byDay[slot.day] = byDay[slot.day] || []).push({ rec, slot });
    });
  });

  // Un mismo NRC puede tener varias filas para el mismo bloque (tramos de
  // fechas con distinta sala). Se agrupan en un único tope por par de bloques.
  const conflicts = [];
  const byKey = new Map();
  const blockKey = (rec, slot) => `${rec.nrc}|${rec.tipo}|${slot.range.text}`;

  Object.keys(byDay).forEach(day => {
    const list = byDay[day];
    for (let i = 0; i < list.length; i++) {
      for (let j = i + 1; j < list.length; j++) {
        const A = list[i], B = list[j];
        const a = A.rec, b = B.rec;
        if (!timesOverlap(A.slot.range, B.slot.range)) continue;
        if (shouldSkipPair(a, b)) continue;

        let kind = '';
        let when = '';
        let datePart = '';
        let ovStart = null;
        let ovEnd = null;

        if (a.weekly && b.weekly) {
          if (!periodsOverlap(a, b)) continue;
          kind = 'Clase / clase';
          when = `${DAY_LABEL[day]} ${A.slot.range.text} vs ${B.slot.range.text}`;
          if (a.start != null && b.start != null) {
            ovStart = Math.max(a.start, b.start);
            ovEnd = Math.min(a.end, b.end);
          }
        } else if (a.eval && b.eval) {
          if (a.start == null || a.start !== b.start) continue;
          kind = 'Evaluación / evaluación';
          datePart = a.meta.inicio;
          when = `${a.meta.inicio} · ${A.slot.range.text} vs ${B.slot.range.text}`;
        } else {
          const ev = a.eval ? A : B;
          const wk = a.eval ? B : A;
          if (!dateInside(ev.rec.start, wk.rec)) continue;
          if (evalInsideOwnSchedule(ev.rec, ev.slot)) continue;
          kind = 'Evaluación / clase';
          datePart = ev.rec.meta.inicio;
          when = `${ev.rec.meta.inicio} (${DAY_LABEL[day]}) · ${ev.slot.range.text} vs ${wk.slot.range.text}`;
        }

        // Orden estable del par: A siempre es el de menor clave de bloque.
        let first = A, second = B;
        if (blockKey(B.rec, B.slot) < blockKey(A.rec, A.slot)) { first = B; second = A; }
        const key = [kind, day, datePart, blockKey(first.rec, first.slot), blockKey(second.rec, second.slot)].join('#');

        let c = byKey.get(key);
        if (!c) {
          c = { kind, when, a: first.rec, b: second.rec, day, rows: new Set(), ovStart, ovEnd };
          byKey.set(key, c);
          conflicts.push(c);
        } else if (ovStart != null) {
          c.ovStart = Math.min(c.ovStart, ovStart);
          c.ovEnd = Math.max(c.ovEnd, ovEnd);
        }
        c.rows.add(a.index).add(b.index);
        // Registros de todas las filas involucradas, para marcar cada una.
        (c.recs = c.recs || new Map()).set(a.index, a).set(b.index, b);
      }
    }
  });

  conflicts.forEach(c => {
    if (c.kind === 'Clase / clase' && c.ovStart != null) {
      const ddmm = t => { const d = new Date(t); return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}`; };
      c.when += ` (entre ${ddmm(c.ovStart)} y ${ddmm(c.ovEnd)})`;
    }
  });

  const order = { 'Clase / clase': 0, 'Evaluación / clase': 1, 'Evaluación / evaluación': 2 };
  conflicts.sort((x, y) => (order[x.kind] - order[y.kind]) || (x.a.index - y.a.index));
  return conflicts;
}

function describeRow(rec) {
  const m = rec.meta;
  return `NRC ${m.nrc} · ${m.materia}${m.curso} ${m.titulo} · ${m.tipo}`;
}

/* --------------------------------------------------------------------------
 * Normativa de topes y bloqueos académicos
 *
 * Criterios de excepciones de la Facultad: se puede solicitar tope solo entre
 * AYUD-AYUD, CLAS-AYUD, LAB-AYUD, LAB-CLAS y LAB-LAB. Nunca CLAS-CLAS ni topes
 * que involucren evaluaciones. Solo se permite un tope por alumno.
 *
 * Bloqueo académico (data/bloqueos.json, hoja BLOQUEOS): un bloque está
 * bloqueado si está marcado BLOQUEADO (amarillo), o si es AYUD con
 * EVALUACIONES (azul) y NUM EVALUACIONES >= umbral (4). Un tope permitido se
 * autoriza si al menos uno de los dos bloques está libre; si ambos están
 * bloqueados → BLOQUEO ACADÉMICO.
 * ------------------------------------------------------------------------ */

const TIPO_SHORT = { CLAS: 'CLAS', AYUD: 'AYUD', 'LAB/TALLER': 'LAB', PRBA: 'PRBA', EXAM: 'EXAM' };
const STATUS_STYLE = {
  NO_PERMITIDO: { label: 'No permitido', cls: 'st-no', cell: '#fecaca' },
  BLOQUEO: { label: 'Bloqueo académico', cls: 'st-bloqueo', cell: '#fecaca' },
  AUTORIZABLE: { label: 'Autorizable con solicitud', cls: 'st-ok', cell: '#fde68a' }
};

function bloqueoEntry(meta) {
  const cursos = BLOQUEOS?.cursos || {};
  const code = courseCode(meta);
  if (cursos[code]) return { code, entry: cursos[code], via: '' };
  for (const k of [meta.lc, meta.conector]) {
    const inherited = LC_BLOQUEO_CODE[normalizeText(k)];
    if (inherited) return { code: inherited, entry: cursos[inherited], via: normalizeText(k) };
  }
  return null; // sin información → se considera libre
}

function isBlockBlocked(rec) {
  const info = bloqueoEntry(rec.meta);
  if (!info) return { blocked: false, info: null, reason: 'sin restricción registrada' };
  const key = TIPO_SHORT[rec.tipo];
  const mark = info.entry[key] || '';
  const minEval = Number(STATIC_CONFIG?.bloqueoAyudMinEvaluaciones || 4);
  if (mark === 'B') return { blocked: true, info, reason: `${key} bloqueada` };
  if (key === 'AYUD' && mark === 'E' && Number(info.entry.numEval || 0) >= minEval) {
    return { blocked: true, info, reason: `AYUD con ${info.entry.numEval} evaluaciones` };
  }
  return { blocked: false, info, reason: `${key} libre` };
}

function describeBlock(rec, st) {
  const code = courseCode(rec.meta);
  const src = st.info && st.info.code !== code ? ` (vía LC ${st.info.via} → ${st.info.code})` : '';
  return `${code} ${st.reason}${src}`;
}

function classifyConflict(c) {
  const tA = TIPO_SHORT[c.a.tipo] || c.a.tipo;
  const tB = TIPO_SHORT[c.b.tipo] || c.b.tipo;
  const pair = [tA, tB].sort().join('-');

  if (c.a.eval || c.b.eval) {
    return { status: 'NO_PERMITIDO', pair, detail: 'Involucra una evaluación' };
  }
  if (tA === 'CLAS' && tB === 'CLAS') {
    return { status: 'NO_PERMITIDO', pair, detail: 'No se aceptan topes entre clases' };
  }

  const sA = isBlockBlocked(c.a);
  const sB = isBlockBlocked(c.b);
  const detail = `${describeBlock(c.a, sA)} · ${describeBlock(c.b, sB)}`;
  if (sA.blocked && sB.blocked) return { status: 'BLOQUEO', pair, detail };
  return { status: 'AUTORIZABLE', pair, detail };
}

function markCell(rowIndex, colIndex, message, color = '#fecaca') {
  const tr = currentRowElements[rowIndex];
  const td = tr?.children[colIndex];
  if (!td) return;
  // Rojo tiene prioridad sobre ámbar si una celda participa en varios topes.
  if (td.dataset.mark !== '#fecaca') {
    td.style.background = color;
    td.dataset.mark = color;
  }
  td.classList.add('cell-conflict');
  const prev = td.title ? td.title.split('\n') : [];
  if (!prev.includes(message)) prev.push(message);
  td.title = prev.join('\n');
}

function clearConflictHighlights() {
  currentRowElements.forEach(tr => {
    [...tr.children].forEach(td => {
      td.style.background = td.dataset.baseBg || '#ffffff';
      td.classList.remove('cell-conflict');
      delete td.dataset.mark;
      td.title = '';
    });
  });
}

function renderConflictList(conflicts, mode) {
  const card = document.getElementById('conflictCard');
  const body = document.getElementById('conflictBody');
  const title = document.getElementById('conflictTitle');
  const withNorm = mode === 'bloqueos';
  body.innerHTML = '';
  document.getElementById('conflictNormHeader').classList.toggle('hidden', !withNorm);
  document.getElementById('conflictNote').classList.toggle('hidden', !withNorm);

  if (withNorm) {
    const n = { NO_PERMITIDO: 0, BLOQUEO: 0, AUTORIZABLE: 0 };
    conflicts.forEach(c => n[c.norm.status]++);
    // "Solo se permite un tope de horario por alumno": se cuentan pares de cursos.
    const autorizablePairs = new Set(conflicts
      .filter(c => c.norm.status === 'AUTORIZABLE')
      .map(c => [c.a.nrc, c.b.nrc].sort().join('|')));
    let t = `Topes encontrados: ${conflicts.length} · No permitidos: ${n.NO_PERMITIDO} · Bloqueo académico: ${n.BLOQUEO} · Autorizables: ${n.AUTORIZABLE}`;
    if (studentContext && autorizablePairs.size > 1) {
      t += ` · ⚠️ ${autorizablePairs.size} pares de cursos con tope autorizable: la normativa permite solo un tope por alumno`;
    }
    title.textContent = t;
  } else {
    title.textContent = `Topes encontrados: ${conflicts.length}`;
  }

  conflicts.forEach(c => {
    const tr = document.createElement('tr');
    const kind = withNorm ? `${c.norm.pair}` : c.kind;
    let html = `
      <td class="conflict-kind">${escapeHtml(kind)}</td>
      <td>${escapeHtml(c.when)}</td>
      <td>${escapeHtml(describeRow(c.a))}</td>
      <td>${escapeHtml(describeRow(c.b))}</td>`;
    if (withNorm) {
      const st = STATUS_STYLE[c.norm.status];
      html += `<td><span class="status-badge ${st.cls}">${escapeHtml(st.label)}</span>` +
              `<div class="status-detail">${escapeHtml(c.norm.detail)}</div></td>`;
    }
    tr.innerHTML = html;
    tr.addEventListener('click', () => scrollToRow(c.a.index));
    body.appendChild(tr);
  });

  card.classList.toggle('hidden', conflicts.length === 0);
}

function scrollToRow(rowIndex) {
  const tr = currentRowElements[rowIndex];
  if (!tr) return;
  tr.scrollIntoView({ behavior: 'smooth', block: 'center' });
  tr.classList.remove('row-flash');
  void tr.offsetWidth;
  tr.classList.add('row-flash');
}

function runConflictReview({ silent = false, mode = 'basic' } = {}) {
  clearConflictHighlights();
  const conflicts = detectConflicts(currentRenderedRows);
  const withNorm = mode === 'bloqueos';

  conflicts.forEach(c => {
    let msg = `${c.kind}: NRC ${c.a.meta.nrc} y NRC ${c.b.meta.nrc}`;
    let color = '#fecaca';
    if (withNorm) {
      c.norm = classifyConflict(c);
      const st = STATUS_STYLE[c.norm.status];
      msg = `${c.norm.pair} · ${st.label}: NRC ${c.a.meta.nrc} y NRC ${c.b.meta.nrc}`;
      color = st.cell;
    }
    c.recs.forEach(rec => {
      markCell(rec.index, DAY_COLUMN_INDEX[c.day], msg, color);
      if (rec.eval) markCell(rec.index, INICIO_COLUMN_INDEX, msg, color);
    });
  });

  if (withNorm) {
    const order = { BLOQUEO: 0, NO_PERMITIDO: 1, AUTORIZABLE: 2 };
    conflicts.sort((x, y) => order[x.norm.status] - order[y.norm.status] || x.a.index - y.a.index);
  }

  conflictReviewActive = true;
  reviewMode = mode;
  updateConflictButton();
  renderConflictList(conflicts, mode);

  if (!silent && !conflicts.length) {
    showModal('Revisión de topes', 'No se encuentran topes en la selección actual.');
  }
  return conflicts;
}

function startReview(mode) {
  if (conflictReviewActive && reviewMode === mode) {
    resetConflictReviewState();
    return;
  }
  if (mode === 'bloqueos' && !BLOQUEOS?.cursos) {
    showModal('Revisión de bloqueos', 'Aún no se ha publicado la información de bloqueos académicos (data/bloqueos.json).');
    return;
  }
  if (currentRenderedRows.length > 1500) {
    showModal('Revisión de topes', 'La selección es muy grande. Filtra primero (por ejemplo por NRC o cargando un alumno) para revisar topes.');
    return;
  }
  runConflictReview({ mode });
}

function toggleConflictReview() { startReview('basic'); }
function toggleBloqueoReview() { startReview('bloqueos'); }

function resetConflictReviewState() {
  conflictReviewActive = false;
  updateConflictButton();
  clearConflictHighlights();
  document.getElementById('conflictCard').classList.add('hidden');
}

function updateConflictButton() {
  document.getElementById('conflictBtn')?.classList.toggle('active', conflictReviewActive && reviewMode === 'basic');
  document.getElementById('bloqueoBtn')?.classList.toggle('active', conflictReviewActive && reviewMode === 'bloqueos');
}

/* --------------------------------------------------------------------------
 * Alumno: cargar NRC inscritos por RUT (consulta al Web App)
 * ------------------------------------------------------------------------ */

function setStudentStatus(html, isError = false) {
  const el = document.getElementById('studentStatus');
  el.innerHTML = html;
  el.classList.toggle('is-error', isError);
  el.classList.toggle('hidden', !html);
}

function setStudentBusy(busy) {
  const btn = document.getElementById('btnBuscarAlumno');
  btn.disabled = busy;
  btn.textContent = busy ? 'Consultando...' : 'Cargar ramos inscritos';
}

async function fetchInscripciones(rutNorm) {
  const base = String(STATIC_CONFIG?.rutLookupUrl || '').trim();
  if (!base) throw new Error('La consulta por RUT no está configurada (falta rutLookupUrl en config.json).');

  const url = `${base}${base.includes('?') ? '&' : '?'}action=nrcsPorRut&rut=${encodeURIComponent(rutNorm)}`;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), RUT_LOOKUP_TIMEOUT_MS);
  try {
    const res = await fetch(url, { signal: controller.signal });
    if (!res.ok) throw new Error(`El servicio respondió ${res.status}.`);
    const data = await res.json();
    if (!data || data.ok === false) throw new Error(data?.error || 'Respuesta inválida del servicio.');
    return data;
  } catch (err) {
    if (err.name === 'AbortError') throw new Error('El servicio tardó demasiado en responder. Intenta nuevamente.');
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

async function buscarAlumno() {
  const input = document.getElementById('studentRut');
  const rut = normalizeRut(input.value);

  if (!isValidRut(rut)) {
    setStudentStatus('RUT no válido. Revisa los dígitos y el dígito verificador.', true);
    return;
  }

  setStudentBusy(true);
  setStudentStatus('');
  try {
    const data = await fetchInscripciones(rut);
    applyStudent(rut, data);
  } catch (err) {
    setStudentStatus(`No se pudieron obtener los ramos inscritos: ${escapeHtml(err.message)}`, true);
  } finally {
    setStudentBusy(false);
  }
}

function applyStudent(rut, data) {
  const inscripciones = Array.isArray(data.inscripciones) ? data.inscripciones : [];
  const rutFmt = formatRut(rut);

  if (!inscripciones.length) {
    // No dejar aplicados los filtros de un alumno consultado antes.
    if (studentContext) limpiar();
    studentContext = null;
    document.getElementById('btnLimpiarAlumno').classList.add('hidden');
    setStudentStatus(`No se encontraron ramos inscritos para el RUT ${escapeHtml(rutFmt)}${data.periodo ? ` en el periodo ${escapeHtml(data.periodo)}` : ''}.`, true);
    return;
  }

  // Valores NRC tal como aparecen en data.json (para que coincidan con el filtro).
  const nrcDisplay = {};
  (STATIC_DATA?.rows || []).forEach(r => { nrcDisplay[normalizeNrc(r.meta.nrc)] = r.meta.nrc; });

  const conHorario = [];
  const sinHorario = [];
  const seen = new Set();
  inscripciones.forEach(ins => {
    const key = normalizeNrc(ins.nrc);
    if (!key || seen.has(key)) return;
    seen.add(key);
    if (DATA_NRC_SET.has(key)) conHorario.push({ ...ins, display: nrcDisplay[key] });
    else sinHorario.push(ins);
  });

  studentContext = { rut, conHorario, sinHorario };

  resetAllFilters();
  selectedValuesState.nrc = conHorario.map(x => x.display);
  resetConflictReviewState();
  queryAndRender();

  const conflicts = conHorario.length > 1 ? runConflictReview({ silent: true }) : [];

  let html = `<strong>Alumno ${escapeHtml(rutFmt)}</strong>: ${seen.size} NRC inscritos`;
  if (data.periodo) html += ` (periodo ${escapeHtml(data.periodo)})`;
  html += ` · ${conHorario.length} con horario en esta planilla.`;
  html += conflicts.length
    ? ` <strong>${conflicts.length} tope(s) detectado(s)</strong> — ver detalle abajo.`
    : (conHorario.length > 1 ? ' Sin topes detectados.' : '');

  if (sinHorario.length) {
    html += `<div class="muted">Inscritos sin horario en esta planilla (no se revisan topes):</div><ul>` +
      sinHorario.map(s => `<li>NRC ${escapeHtml(s.nrc)} · ${escapeHtml(`${s.materia || ''}${s.curso || ''}`)} ${escapeHtml(s.asignatura || '')}</li>`).join('') +
      '</ul>';
  }

  setStudentStatus(html);
  document.getElementById('btnLimpiarAlumno').classList.remove('hidden');
}

function limpiarAlumno() {
  studentContext = null;
  document.getElementById('studentRut').value = '';
  document.getElementById('btnLimpiarAlumno').classList.add('hidden');
  setStudentStatus('');
  limpiar();
}

/* --------------------------------------------------------------------------
 * Otros
 * ------------------------------------------------------------------------ */

function showModal(title, message) {
  document.getElementById('modalTitle').textContent = title;
  document.getElementById('modalMessage').textContent = message;
  document.getElementById('modalOverlay').classList.remove('hidden');
}

function closeModal() {
  document.getElementById('modalOverlay').classList.add('hidden');
}

function descargarExcelOriginal() {
  const url = STATIC_CONFIG?.horarioOnlyDownloadUrl || '';
  if (!url) {
    showModal('Error', 'No se encontró la URL de descarga de HORARIO en config.json.');
    return;
  }
  window.open(url, '_blank', 'noopener');
}

function syncScrollWidth() {
  const table = document.getElementById('mainTable');
  const inner = document.getElementById('topScrollInner');
  if (table && inner) inner.style.width = table.scrollWidth + 'px';
}

function initScrollSync() {
  const top = document.getElementById('topScroll');
  const wrap = document.getElementById('tableWrap');
  let lock = false;
  top.addEventListener('scroll', () => { if (lock) return; lock = true; wrap.scrollLeft = top.scrollLeft; lock = false; });
  wrap.addEventListener('scroll', () => { if (lock) return; lock = true; top.scrollLeft = wrap.scrollLeft; lock = false; });
  window.addEventListener('resize', syncScrollWidth);
}

window.addEventListener('load', async () => {
  try {
    setLoading(true);
    await loadStaticFiles();
    renderPeriodMeta();
    initScrollSync();
    buildFilters();
    updateDayFilterButtons();
    queryAndRender();
    renderLegend(STATIC_DATA?.legendRows || []);
  } catch (err) {
    showModal('Error', 'No se pudieron cargar los archivos del horario: ' + err.message);
  } finally {
    setLoading(false);
  }
});
