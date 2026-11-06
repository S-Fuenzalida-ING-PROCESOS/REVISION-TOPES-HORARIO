/**
 * REVISION-TOPES-HORARIO · Apps Script (archivo Web.gs)
 *
 * Este archivo reemplaza TODO lo que en el script anterior estaba ANTES de
 * `function normalizeName(str)`. La lógica de construcción del horario
 * (normalizeName … updateSalasPruebas) va sin cambios en Horario.gs.
 *
 * Qué hace este archivo:
 *  - Parámetros del periodo (hoja PARAMETROS_HORARIO).
 *  - Importa el reporte UG307 a una hoja INSCRIPCIONES minimizada
 *    (solo RUT normalizado + NRC + curso; sin nombres).
 *  - Web App:
 *      ?action=nrcsPorRut&rut=123456789  → JSON con NRC inscritos de ese RUT
 *      ?action=downloadHorarioExcel      → descarga de la hoja HORARIO
 *  - Publica data/config.json y data/data.json en GitHub.
 *
 * El token de GitHub SOLO se guarda en Script Properties (menú
 * "Configurar GitHub"). No se lee desde ninguna hoja.
 */

const SOURCE_SS_ID = '';  // Sin valor por defecto: usar "Configurar MAESTRO fuente".

const PARAMS_SHEET_NAME = 'PARAMETROS_HORARIO';
const DEFAULT_PERIODO = '202710';
const DEFAULT_BRANCH = 'main';
const DEFAULT_GITHUB_OWNER = 'S-Fuenzalida-ING-PROCESOS';
const DEFAULT_GITHUB_REPO = 'REVISION-TOPES-HORARIO';

const SOURCE_SS_ID_PROP = 'SOURCE_SS_ID';
const SOURCE_PROFESORES_SHEET_NAME = 'PROFESORES';
const LOCAL_PROFESORES_ING_SHEET_NAME = 'PROFESORES_ING';

const UG307_RAW_SHEET_NAME = 'UG307';
const INSCRIPCIONES_SHEET_NAME = 'INSCRIPCIONES';
const INSCRIPCIONES_HEADER = ['RUT', 'NRC', 'MATERIA', 'CURSO', 'ASIGNATURA', 'PERIODO'];
const INSCRIPCIONES_VERSION_PROP = 'INSCRIPCIONES_VERSION';
const RUT_LOOKUP_CACHE_SECONDS = 600;
const RUT_LOOKUP_MAX_PER_MINUTE = 240;

const WEBAPP_URL_PROP = 'HORARIO_DOWNLOAD_WEBAPP_URL';
const HORARIO_ONLY_DOWNLOAD_FILE_ID_PROP = 'HORARIO_ONLY_DOWNLOAD_FILE_ID';
const HORARIO_EXPORT_SOURCE_VERSION_PROP = 'HORARIO_EXPORT_SOURCE_VERSION';
const DOWNLOAD_LOCK_TIMEOUT_MS = 120000;
const DOWNLOAD_EXPORT_SETTLE_MS = 1500;

/*******************************
 * PARAMETROS
 *******************************/
function ensureParametrosHorarioSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PARAMS_SHEET_NAME);

  const defaultRows = [
    ['KEY', 'VALUE', 'DESCRIPCION'],
    ['PERIODO', DEFAULT_PERIODO, 'Periodo académico (ej: 202710)'],
    ['CLASS_START', '', 'Inicio de CLAS (fecha)'],
    ['CLASS_END', '', 'Fin de CLAS (fecha)'],
    ['AYUD_START', '', 'Inicio de AYUD/LAB (fecha)'],
    ['AYUD_END', '', 'Fin de AYUD/LAB (fecha)']
  ];

  if (!sh) {
    sh = ss.insertSheet(PARAMS_SHEET_NAME);
    sh.getRange(1, 1, defaultRows.length, 3).setValues(defaultRows);
    sh.getRange('A1:C1').setFontWeight('bold');
    sh.autoResizeColumns(1, 3);
    return sh;
  }

  const existing = sh.getDataRange().getValues();
  if (!existing.length || String(existing[0][0] || '').trim().toUpperCase() !== 'KEY') {
    sh.insertRows(1, defaultRows.length);
    sh.getRange(1, 1, defaultRows.length, 3).setValues(defaultRows);
    sh.getRange('A1:C1').setFontWeight('bold');
    return sh;
  }

  const keys = new Set(existing.slice(1).map(r => String(r[0] || '').trim().toUpperCase()));
  const missing = defaultRows.slice(1).filter(r => !keys.has(r[0]));
  if (missing.length) {
    sh.getRange(sh.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
  }
  return sh;
}

function getParamsMap_() {
  const sh = ensureParametrosHorarioSheet_();
  const values = sh.getDataRange().getValues();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim().toUpperCase();
    if (!key) continue;
    const raw = values[i][1];
    // Corrección: antes se usaba String(fecha), que producía
    // "Mon Aug 03 2026 00:00:00 GMT-0400" en vez de DD/MM/YYYY.
    map[key] = (raw instanceof Date && !isNaN(raw))
      ? Utilities.formatDate(raw, tz, 'dd/MM/yyyy')
      : String(raw == null ? '' : raw).trim();
  }
  return map;
}

function getHorarioParams_() {
  const map = getParamsMap_();
  return {
    classStart: map['CLASS_START'] || '',
    classEnd: map['CLASS_END'] || '',
    ayudStart: map['AYUD_START'] || '',
    ayudEnd: map['AYUD_END'] || ''
  };
}

function getPeriodo_() {
  return getParamsMap_()['PERIODO'] || DEFAULT_PERIODO;
}

function getHorarioDownloadFileName_() {
  return 'HORARIO_ING_' + getPeriodo_();
}

/*******************************
 * MAESTRO FUENTE
 *******************************/
function getSourceSpreadsheetId_() {
  const fromProps = (PropertiesService.getScriptProperties().getProperty(SOURCE_SS_ID_PROP) || '').trim();
  const id = fromProps || SOURCE_SS_ID;
  if (!id) throw new Error('Falta configurar el MAESTRO fuente (menú: Configurar MAESTRO fuente).');
  return id;
}

function setupSourceMaestroId() {
  const ui = SpreadsheetApp.getUi();
  let current = '';
  try { current = getSourceSpreadsheetId_(); } catch (e) { current = '(no configurado)'; }

  const resp = ui.prompt(
    'Configurar MAESTRO fuente ' + getPeriodo_(),
    'Pega el ID del Google Spreadsheet que contiene las hojas MAESTRO y PROFESORES.\n\nID actual:\n' + current,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const id = String(resp.getResponseText() || '').trim();
  if (!id) throw new Error('Debes ingresar un ID de spreadsheet válido.');

  const src = SpreadsheetApp.openById(id);
  if (!src.getSheetByName('MAESTRO')) throw new Error('El archivo indicado no contiene una hoja MAESTRO.');
  if (!src.getSheetByName(SOURCE_PROFESORES_SHEET_NAME)) {
    throw new Error('El archivo indicado no contiene una hoja ' + SOURCE_PROFESORES_SHEET_NAME + '.');
  }

  PropertiesService.getScriptProperties().setProperty(SOURCE_SS_ID_PROP, id);
  ui.alert('MAESTRO fuente configurado.\n\nArchivo: ' + src.getName());
}

function showSourceMaestroId() {
  let msg;
  try {
    const id = getSourceSpreadsheetId_();
    msg = 'ID MAESTRO fuente actual:\n\n' + id + '\n\nArchivo: ' + SpreadsheetApp.openById(id).getName();
  } catch (e) {
    msg = e.message;
  }
  SpreadsheetApp.getUi().alert(msg);
}

/*******************************
 * INSCRIPCIONES (UG307)
 *******************************/

/**
 * Pasos para el usuario:
 *  1. Archivo > Importar > subir el Excel UG307 > "Insertar hojas nuevas".
 *  2. Renombrar la hoja importada a "UG307".
 *  3. Menú Macros > Importar inscripciones UG307.
 * Se genera la hoja INSCRIPCIONES (sin nombres de alumnos) y se ofrece
 * eliminar la hoja UG307 cruda.
 */
function importarInscripcionesUG307() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName(UG307_RAW_SHEET_NAME);
  if (!raw) {
    ui.alert(
      'No encontré la hoja "' + UG307_RAW_SHEET_NAME + '".\n\n' +
      'Importa el Excel UG307 (Archivo > Importar > Insertar hojas nuevas) y renombra la hoja a "' +
      UG307_RAW_SHEET_NAME + '".'
    );
    return;
  }

  const values = raw.getDataRange().getDisplayValues();
  if (values.length < 2) {
    ui.alert('La hoja UG307 no tiene datos.');
    return;
  }

  const idx = buildHeaderIndexNormalized_(values[0]);
  const col = {
    rut: getFirstExistingIndex_(idx, ['RUT']),
    nrc: getFirstExistingIndex_(idx, ['NRC']),
    materia: getFirstExistingIndex_(idx, ['MATERIA']),
    curso: getFirstExistingIndex_(idx, ['CURSO']),
    asignatura: getFirstExistingIndex_(idx, ['ASIGNATURA', 'TITULO']),
    periodo: getFirstExistingIndex_(idx, ['PERIODO INSCRIPCION', 'PERIODO INSCRIPCIÓN', 'PERIODO'])
  };
  for (const k of ['rut', 'nrc']) {
    if (col[k] < 0) {
      ui.alert('Falta la columna ' + k.toUpperCase() + ' en la hoja UG307.');
      return;
    }
  }

  const periodo = getPeriodo_();
  const seen = new Set();
  const out = [];
  let otrosPeriodos = 0;
  let rutInvalidos = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rut = normalizeRut_(row[col.rut]);
    const nrc = normalizeNrcKey_(row[col.nrc]);
    if (!rut || !nrc) continue;
    if (!/^\d{7,8}[0-9K]$/.test(rut)) { rutInvalidos++; continue; }

    const per = col.periodo >= 0 ? String(row[col.periodo] || '').trim() : '';
    if (per && periodo && per !== periodo) { otrosPeriodos++; continue; }

    const key = rut + '|' + nrc;
    if (seen.has(key)) continue;
    seen.add(key);

    out.push([
      rut,
      nrc,
      col.materia >= 0 ? String(row[col.materia] || '').trim() : '',
      col.curso >= 0 ? String(row[col.curso] || '').trim() : '',
      col.asignatura >= 0 ? String(row[col.asignatura] || '').trim() : '',
      per
    ]);
  }

  if (!out.length) {
    ui.alert(
      'No quedaron inscripciones para importar.\n\n' +
      'Filas de otro periodo (distinto de ' + periodo + '): ' + otrosPeriodos + '\n' +
      'Revisa el parámetro PERIODO en ' + PARAMS_SHEET_NAME + '.'
    );
    return;
  }

  out.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : (a[1] < b[1] ? -1 : 1));

  let sh = ss.getSheetByName(INSCRIPCIONES_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(INSCRIPCIONES_SHEET_NAME);
  sh.clear();
  sh.getRange(1, 1, 1, INSCRIPCIONES_HEADER.length).setValues([INSCRIPCIONES_HEADER]).setFontWeight('bold');
  // Todo como texto: conserva ceros a la izquierda en CURSO y RUT con K.
  sh.getRange(2, 1, out.length, INSCRIPCIONES_HEADER.length).setNumberFormat('@').setValues(out);
  sh.setFrozenRows(1);

  // Nueva versión → invalida el caché de consultas por RUT.
  PropertiesService.getScriptProperties().setProperty(INSCRIPCIONES_VERSION_PROP, String(Date.now()));

  const alumnos = new Set(out.map(r => r[0])).size;
  const resp = ui.alert(
    'Inscripciones importadas',
    'Periodo: ' + periodo + '\n' +
    'Alumnos: ' + alumnos + '\n' +
    'Inscripciones (RUT+NRC): ' + out.length + '\n' +
    'Omitidas por otro periodo: ' + otrosPeriodos + '\n' +
    'Omitidas por RUT inválido: ' + rutInvalidos + '\n\n' +
    '¿Eliminar la hoja UG307 original? (contiene nombres de alumnos que la web no necesita)',
    ui.ButtonSet.YES_NO
  );
  if (resp === ui.Button.YES) ss.deleteSheet(raw);
}

function normalizeNrcKey_(value) {
  const s = String(value == null ? '' : value).trim();
  if (/^\d+(\.0+)?$/.test(s)) return String(Number(s));
  return s.toUpperCase();
}

function lookupInscripcionesPorRut_(rutNorm) {
  const props = PropertiesService.getScriptProperties();
  const version = props.getProperty(INSCRIPCIONES_VERSION_PROP) || '0';
  const cache = CacheService.getScriptCache();
  const cacheKey = 'insc_' + version + '_' + rutNorm;

  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INSCRIPCIONES_SHEET_NAME);
  if (!sh) throw new Error('Aún no se han cargado inscripciones para este periodo.');

  const lastRow = sh.getLastRow();
  const result = [];
  if (lastRow >= 2) {
    const matches = sh.getRange(2, 1, lastRow - 1, 1)
      .createTextFinder(rutNorm)
      .matchEntireCell(true)
      .findAll();

    if (matches.length) {
      // INSCRIPCIONES está ordenada por RUT: las coincidencias son contiguas.
      const first = matches[0].getRow();
      const last = matches[matches.length - 1].getRow();
      const rows = sh.getRange(first, 1, last - first + 1, INSCRIPCIONES_HEADER.length).getDisplayValues();
      rows.forEach(r => {
        if (r[0] !== rutNorm) return;
        result.push({ nrc: r[1], materia: r[2], curso: r[3], asignatura: r[4] });
      });
    }
  }

  try { cache.put(cacheKey, JSON.stringify(result), RUT_LOOKUP_CACHE_SECONDS); } catch (e) {}
  return result;
}

function isRateLimited_() {
  const cache = CacheService.getScriptCache();
  const key = 'rl_' + Math.floor(Date.now() / 60000);
  const count = Number(cache.get(key) || '0') + 1;
  cache.put(key, String(count), 120);
  return count > RUT_LOOKUP_MAX_PER_MINUTE;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleNrcsPorRut_(e) {
  try {
    const rut = normalizeRut_(e && e.parameter ? e.parameter.rut : '');
    if (!/^\d{7,8}[0-9K]$/.test(rut)) {
      return jsonOutput_({ ok: false, error: 'RUT inválido.' });
    }
    if (isRateLimited_()) {
      return jsonOutput_({ ok: false, error: 'Demasiadas consultas en este minuto. Intenta nuevamente en unos segundos.' });
    }
    return jsonOutput_({
      ok: true,
      periodo: getPeriodo_(),
      inscripciones: lookupInscripcionesPorRut_(rut)
    });
  } catch (err) {
    return jsonOutput_({ ok: false, error: err.message || String(err) });
  }
}

/*******************************
 * WEB APP
 *******************************/
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : '';

  if (action === 'nrcsPorRut') {
    return handleNrcsPorRut_(e);
  }

  if (action === 'downloadHorarioExcel') {
    try {
      // Solo se regenera si la planilla cambió desde la última exportación.
      const fileId = ensureHorarioDownloadReady_();
      return buildInlineXlsxDownloadHtml_({
        fileId: fileId,
        fileName: getHorarioDownloadFileName_() + '.xlsx',
        message: 'El HORARIO se descargará automáticamente en unos segundos'
      });
    } catch (err) {
      logWebPublish_('DOWNLOAD_HORARIO_EXCEL', 'ERROR', err.message || String(err));
      return buildDownloadErrorHtml_('No se pudo preparar la descarga de HORARIO.', err);
    }
  }

  return HtmlService.createHtmlOutput(
    '<html><body><p>Web App activa.</p>' +
    '<p>?action=downloadHorarioExcel · ?action=nrcsPorRut&amp;rut=…</p></body></html>'
  );
}

function setupWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Configurar URL del Web App',
    'Pega la URL pública del Web App (termina en /exec).\n' +
    'Implementar > Nueva implementación > Aplicación web · Ejecutar como: Yo · Acceso: Cualquier persona.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const url = String(resp.getResponseText() || '').trim();
  if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) {
    throw new Error('Debes pegar una URL válida del Web App que termine en /exec');
  }
  PropertiesService.getScriptProperties().setProperty(WEBAPP_URL_PROP, url);
  ui.alert('URL del Web App guardada.');
}

function getWebAppUrl_() {
  return (PropertiesService.getScriptProperties().getProperty(WEBAPP_URL_PROP) || '').trim();
}

function getHorarioOnlyExcelEndpointUrl_() {
  const base = getWebAppUrl_();
  return base ? base + '?action=downloadHorarioExcel' : '';
}

/*******************************
 * PUBLICACION GITHUB
 *******************************/
function publishAllStaticFilesToGitHub() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const cfg = getGitHubConfig_();
    validateGitHubConfig_(cfg);

    const downloadFileId = refreshHorarioOnlyDownloadFile_(true);
    const configPayload = buildConfigPayload_();
    const dataPayload = buildStaticPayload_();

    validateWebPayloadsBeforePublish_(configPayload, dataPayload, downloadFileId);

    const payloads = {
      'data/config.json': configPayload,
      'data/data.json': dataPayload
    };

    // Bloqueos académicos: opcional. Si la hoja no existe, la web avisa al usar el botón.
    let bloqueosMsg = 'Hoja ' + BLOQUEOS_SHEET_NAME + ' no encontrada: no se publicaron bloqueos.';
    const bloqueosPayload = buildBloqueosPayload_();
    if (bloqueosPayload) {
      payloads['data/bloqueos.json'] = bloqueosPayload;
      bloqueosMsg = 'Cursos con información de bloqueo: ' + Object.keys(bloqueosPayload.cursos).length;
    }

    const results = Object.keys(payloads).map(path =>
      upsertGitHubFile_(cfg, path, JSON.stringify(payloads[path], null, 1), 'Actualiza ' + path)
    );

    logWebPublish_('PUBLISH', 'SUCCESS', JSON.stringify({ results, horarioRows: dataPayload.rows.length }));

    ui.alert([
      'Publicación completada.',
      '',
      'Repositorio: ' + cfg.owner + '/' + cfg.repo + ' (' + cfg.branch + ')',
      'Periodo: ' + configPayload.periodo,
      'Filas HORARIO publicadas: ' + dataPayload.rows.length,
      bloqueosMsg,
      '',
      ...results.map(r => '- ' + r.path + ' (' + r.action + ')')
    ].join('\n'));
    return results;
  } catch (err) {
    logWebPublish_('PUBLISH', 'ERROR', err.message || String(err));
    ui.alert('Error al publicar:\n\n' + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function buildConfigPayload_() {
  return {
    periodo: getPeriodo_(),
    horarioOnlyDownloadUrl: getHorarioOnlyExcelEndpointUrl_(),
    rutLookupUrl: getWebAppUrl_(),
    bloqueoAyudMinEvaluaciones: BLOQUEO_AYUD_MIN_EVAL
  };
}

/*******************************
 * BLOQUEOS ACADEMICOS (hoja BLOQUEOS)
 *
 * Una fila por curso (códigos del plan 2026). Columnas por encabezado:
 * CODIGO, TITULO, Clases, Ayudantías, Laboratorios o Talleres, NUM EVALUACIONES.
 * Celda con texto BLOQUEADO (o fondo amarillo)  → bloque bloqueado ("B").
 * Celda con texto EVALUACIONES (o fondo azul)   → evaluaciones en el bloque ("E").
 * La regla "AYUD con 4+ evaluaciones = bloqueada" se aplica en la web.
 *******************************/
const BLOQUEOS_SHEET_NAME = 'BLOQUEOS';
const BLOQUEO_AYUD_MIN_EVAL = 4;

function classifyBloqueoCell_(text, bg) {
  const t = String(text || '').toUpperCase();
  if (t.indexOf('BLOQUE') >= 0) return 'B';
  if (t.indexOf('EVALUA') >= 0) return 'E';

  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(String(bg || ''));
  if (!m) return '';
  const r = parseInt(m[1], 16), g = parseInt(m[2], 16), b = parseInt(m[3], 16);
  if (r > 200 && g > 200 && b < 140) return 'B';           // amarillo
  if (b > 200 && b > r + 15 && b > g + 5) return 'E';       // azul
  return '';
}

function buildBloqueosPayload_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BLOQUEOS_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return null;

  const range = sh.getDataRange();
  const values = range.getDisplayValues();
  const bgs = range.getBackgrounds();
  const idx = buildHeaderIndexNormalized_(values[0]);

  const col = {
    codigo: getFirstExistingIndex_(idx, ['CODIGO', 'CÓDIGO']),
    titulo: getFirstExistingIndex_(idx, ['TITULO', 'TÍTULO']),
    numEval: getFirstExistingIndex_(idx, ['NUM EVALUACIONES', 'N EVALUACIONES']),
    CLAS: getFirstExistingIndex_(idx, ['CLASES']),
    AYUD: getFirstExistingIndex_(idx, ['AYUDANTÍAS', 'AYUDANTIAS']),
    LAB: getFirstExistingIndex_(idx, ['LABORATORIOS O TALLERES', 'LABORATORIOS'])
  };
  ['codigo', 'CLAS', 'AYUD', 'LAB'].forEach(k => {
    if (col[k] < 0) throw new Error('Falta la columna ' + k + ' en la hoja ' + BLOQUEOS_SHEET_NAME + '.');
  });

  const cursos = {};
  for (let r = 1; r < values.length; r++) {
    const codigo = String(values[r][col.codigo] || '').replace(/\s+/g, '').toUpperCase();
    if (!codigo) continue;

    const e = cursos[codigo] || (cursos[codigo] = {
      titulo: col.titulo >= 0 ? String(values[r][col.titulo] || '').trim() : '',
      numEval: 0, CLAS: '', AYUD: '', LAB: ''
    });

    const n = col.numEval >= 0 ? parseInt(values[r][col.numEval], 10) : 0;
    if (!isNaN(n)) e.numEval = Math.max(e.numEval, n);

    // Filas duplicadas del mismo código: BLOQUEADO prevalece sobre EVALUACIONES.
    ['CLAS', 'AYUD', 'LAB'].forEach(k => {
      const v = classifyBloqueoCell_(values[r][col[k]], bgs[r][col[k]]);
      if (v === 'B' || (v && !e[k])) e[k] = v;
    });
  }

  return { generatedAt: new Date().toISOString(), periodo: getPeriodo_(), cursos };
}

function validateWebPayloadsBeforePublish_(configPayload, dataPayload, downloadFileId) {
  if (!dataPayload || !Array.isArray(dataPayload.headers) || !dataPayload.headers.length) {
    throw new Error('HORARIO no tiene encabezados válidos.');
  }
  if (!Array.isArray(dataPayload.rows) || !dataPayload.rows.length) {
    throw new Error('HORARIO no tiene filas. Ejecuta Crear Horario, Llenar Salas Horario y Actualizar Salas Pruebas.');
  }
  if (!downloadFileId) throw new Error('No se pudo crear el archivo Excel de descarga.');
  if (!configPayload.rutLookupUrl) {
    throw new Error('Falta la URL del Web App (menú: Configurar URL Web App). Sin ella no funcionan la descarga ni la búsqueda por RUT.');
  }
}

function setupGitHubPublishing() {
  const ui = SpreadsheetApp.getUi();
  const ask = (title, hint, def) => {
    const r = ui.prompt('Configurar GitHub', title + '\n' + hint, ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return null;
    return r.getResponseText().trim() || def || '';
  };

  const owner = ask('OWNER de GitHub', 'Vacío = ' + DEFAULT_GITHUB_OWNER, DEFAULT_GITHUB_OWNER); if (owner === null) return;
  const repo = ask('Repositorio', 'Vacío = ' + DEFAULT_GITHUB_REPO, DEFAULT_GITHUB_REPO); if (repo === null) return;
  const branch = ask('Rama', 'Vacío = ' + DEFAULT_BRANCH, DEFAULT_BRANCH); if (branch === null) return;
  const token = ask(
    'Personal Access Token (fine-grained)',
    'Permiso: Contents = Read and write, solo sobre este repositorio. Se guarda en Script Properties.',
    ''
  ); if (token === null) return;

  if (!token) throw new Error('El token es obligatorio.');

  PropertiesService.getScriptProperties().setProperties({
    GITHUB_OWNER: owner, GITHUB_REPO: repo, GITHUB_BRANCH: branch, GITHUB_TOKEN: token
  }, false);

  ui.alert('Configuración guardada.\n\nOwner: ' + owner + '\nRepo: ' + repo + '\nRama: ' + branch);
}

function showGitHubConfigSummary() {
  const cfg = getGitHubConfig_();
  SpreadsheetApp.getUi().alert(
    'Configuración GitHub:\n\n' +
    'Owner: ' + cfg.owner + '\nRepo: ' + cfg.repo + '\nRama: ' + cfg.branch + '\n' +
    'Token: ' + (cfg.token ? 'guardado' : 'NO guardado') + '\n\n' +
    'URL Web App: ' + (getWebAppUrl_() || '(no configurada)')
  );
}

function testGitHubConnection() {
  const cfg = getGitHubConfig_();
  validateGitHubConfig_(cfg);
  const res = UrlFetchApp.fetch(
    'https://api.github.com/repos/' + encodeURIComponent(cfg.owner) + '/' + encodeURIComponent(cfg.repo),
    { method: 'get', headers: buildGitHubHeaders_(cfg.token), muteHttpExceptions: true }
  );
  const code = res.getResponseCode();
  if (code >= 200 && code < 300) {
    const body = safeJsonParse_(res.getContentText());
    SpreadsheetApp.getUi().alert('Conexión OK.\n\nRepositorio: ' + body.full_name);
    return;
  }
  throw new Error('GitHub respondió ' + code + '.\n\n' + res.getContentText());
}

function getGitHubConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    owner: (props.getProperty('GITHUB_OWNER') || DEFAULT_GITHUB_OWNER).trim(),
    repo: (props.getProperty('GITHUB_REPO') || DEFAULT_GITHUB_REPO).trim(),
    branch: (props.getProperty('GITHUB_BRANCH') || DEFAULT_BRANCH).trim(),
    token: (props.getProperty('GITHUB_TOKEN') || '').trim()
  };
}

function validateGitHubConfig_(cfg) {
  if (!cfg.owner) throw new Error('Falta GITHUB_OWNER.');
  if (!cfg.repo) throw new Error('Falta GITHUB_REPO.');
  if (!cfg.branch) throw new Error('Falta GITHUB_BRANCH.');
  if (!cfg.token) throw new Error('Falta GITHUB_TOKEN (menú: Configurar GitHub).');
}

function upsertGitHubFile_(cfg, path, contentString, commitMessage) {
  const existing = getGitHubFileIfExists_(cfg, path);
  const url = 'https://api.github.com/repos/' + encodeURIComponent(cfg.owner) + '/' +
              encodeURIComponent(cfg.repo) + '/contents/' + encodeURIComponentPath_(path);

  const payload = {
    message: commitMessage,
    content: Utilities.base64Encode(contentString, Utilities.Charset.UTF_8),
    branch: cfg.branch
  };
  if (existing.exists) payload.sha = existing.sha;

  const res = UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    headers: buildGitHubHeaders_(cfg.token),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('No se pudo actualizar ' + path + '. GitHub respondió ' + code + '.\n\n' + res.getContentText());
  }
  return { path, action: existing.exists ? 'updated' : 'created' };
}

function getGitHubFileIfExists_(cfg, path) {
  const url = 'https://api.github.com/repos/' + encodeURIComponent(cfg.owner) + '/' +
              encodeURIComponent(cfg.repo) + '/contents/' + encodeURIComponentPath_(path) +
              '?ref=' + encodeURIComponent(cfg.branch);
  const res = UrlFetchApp.fetch(url, { method: 'get', headers: buildGitHubHeaders_(cfg.token), muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code === 404) return { exists: false, sha: null };
  if (code < 200 || code >= 300) {
    throw new Error('No se pudo leer ' + path + ' en GitHub. Respuesta ' + code + '.\n\n' + res.getContentText());
  }
  return { exists: true, sha: safeJsonParse_(res.getContentText()).sha || null };
}

function buildGitHubHeaders_(token) {
  return {
    Authorization: 'Bearer ' + token,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };
}

function encodeURIComponentPath_(path) {
  return path.split('/').map(encodeURIComponent).join('/');
}

function safeJsonParse_(text) {
  try { return JSON.parse(text); } catch (e) { return { raw: text }; }
}

/*******************************
 * PAYLOAD data.json
 *******************************/
function valueAt_(row, index) {
  return index >= 0 ? row[index] : '';
}

function formatCell_(v) {
  return v == null ? '' : String(v);
}

function buildStaticPayload_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('HORARIO');
  if (!sh) throw new Error('No se encontró la hoja "HORARIO".');

  const range = sh.getDataRange();
  const displayValues = range.getDisplayValues();
  const rawValues = range.getValues();
  const backgrounds = range.getBackgrounds();
  const fontColors = range.getFontColors();

  const headerRow = findHorarioHeaderRow_(sh);
  if (headerRow < 1) throw new Error('No se encontró la fila de encabezados de HORARIO.');

  const headerIndex = headerRow - 1;
  const headers = displayValues[headerIndex];
  const idx = {};
  headers.forEach((h, i) => idx[String(h || '').trim()] = i);

  const cellsOf = r => displayValues[r].map((v, c) => ({
    value: formatCell_(v), bg: backgrounds[r][c], color: fontColors[r][c]
  }));

  const legendRows = [];
  for (let r = 0; r < headerIndex; r++) {
    if (displayValues[r].some(v => String(v || '').trim() !== '')) legendRows.push({ cells: cellsOf(r) });
  }

  const META = {
    area: 'AREA', plan: 'PLAN DE ESTUDIO', nrc: 'NRC', conector: 'CONECTOR DE LIGA', lc: 'LC',
    materia: 'MATERIA', curso: 'CURSO', seccion: 'SECC.', titulo: 'TITULO',
    lunes: 'Lunes', martes: 'Martes', miercoles: 'Miercoles', jueves: 'Jueves', viernes: 'Viernes',
    inicio: 'INICIO', fin: 'FIN', sala: 'SALA', tipo: 'TIPO DE REUNION', profesor: 'PROFESOR'
  };

  const rows = [];
  for (let r = headerIndex + 1; r < displayValues.length; r++) {
    const rowValues = displayValues[r];
    if (!rowValues.some(v => String(v || '').trim() !== '')) continue;

    const meta = {};
    Object.keys(META).forEach(k => { meta[k] = formatCell_(valueAt_(rowValues, idx[META[k]])); });

    rows.push({
      cells: cellsOf(r),
      meta,
      sortInicio: parseDateForSort_(rawValues[r][idx['INICIO']], rowValues[idx['INICIO']])
    });
  }

  rows.sort((a, b) => a.sortInicio - b.sortInicio);
  rows.forEach(r => delete r.sortInicio);

  return { generatedAt: new Date().toISOString(), headers, legendRows, rows };
}

function parseDateForSort_(rawValue, displayValue) {
  if (rawValue instanceof Date && !isNaN(rawValue)) return rawValue.getTime();
  const m = String(displayValue || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])).getTime();
  return Number.MAX_SAFE_INTEGER;
}

/*******************************
 * DIAGNOSTICO
 *******************************/
function diagnoseWebPublishState() {
  const cfg = getGitHubConfig_();
  const lines = ['Diagnóstico', ''];

  lines.push('Periodo: ' + getPeriodo_());
  lines.push('GitHub: ' + cfg.owner + '/' + cfg.repo + ' (' + cfg.branch + ') · token ' + (cfg.token ? 'OK' : 'FALTA'));
  lines.push('URL Web App: ' + (getWebAppUrl_() || 'FALTA'));

  try {
    const p = buildStaticPayload_();
    lines.push('Filas HORARIO: ' + p.rows.length);
  } catch (e) {
    lines.push('HORARIO: ERROR ' + e.message);
  }

  try {
    const bq = buildBloqueosPayload_();
    if (bq) {
      const c = Object.values(bq.cursos);
      const nB = c.reduce((s, e) => s + ['CLAS', 'AYUD', 'LAB'].filter(k => e[k] === 'B').length, 0);
      lines.push('Bloqueos: ' + c.length + ' cursos, ' + nB + ' bloques bloqueados');
    } else {
      lines.push('Bloqueos: hoja ' + BLOQUEOS_SHEET_NAME + ' no encontrada');
    }
  } catch (e) {
    lines.push('Bloqueos: ERROR ' + e.message);
  }

  const ins = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INSCRIPCIONES_SHEET_NAME);
  if (ins && ins.getLastRow() > 1) {
    const ruts = ins.getRange(2, 1, ins.getLastRow() - 1, 1).getDisplayValues().map(r => r[0]);
    lines.push('Inscripciones: ' + ruts.length + ' (' + new Set(ruts).size + ' alumnos)');
  } else {
    lines.push('Inscripciones: NO CARGADAS (menú: Importar inscripciones UG307)');
  }

  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GIT TOKEN')) {
    lines.push('');
    lines.push('⚠️ Existe una hoja "GIT TOKEN". Elimínala: el token debe vivir solo en Script Properties.');
  }

  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

/*******************************
 * DESCARGA EXCEL HORARIO
 *******************************/
function getDownloadSourceSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('El script debe estar vinculado al Google Sheets del horario.');
  return ss;
}

function getSourceSpreadsheetVersion_(ss) {
  return DriveApp.getFileById(ss.getId()).getLastUpdated().getTime();
}

function isReusableDownloadFile_(fileId, expectedName) {
  if (!fileId) return false;
  try {
    return SpreadsheetApp.openById(fileId).getName() === expectedName;
  } catch (e) {
    return false;
  }
}

function horarioDownloadOptions_() {
  return {
    sourceSheetName: 'HORARIO',
    outputSheetName: 'HORARIO',
    fileIdProp: HORARIO_ONLY_DOWNLOAD_FILE_ID_PROP,
    fileName: getHorarioDownloadFileName_(),
    sourceVersionProp: HORARIO_EXPORT_SOURCE_VERSION_PROP,
    filterHeaderRow: 15
  };
}

function cleanExportSheetForDownload_(sheet, headerRow) {
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  if (sheet.getMaxRows() > 0) sheet.showRows(1, sheet.getMaxRows());
  if (sheet.getMaxColumns() > 0) sheet.showColumns(1, sheet.getMaxColumns());

  const header = Number(headerRow || 1);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow >= header && lastCol > 0) {
    sheet.getRange(header, 1, lastRow - header + 1, lastCol).createFilter();
  }
}

function refreshSingleSheetDownloadCore_(options) {
  const ss = getDownloadSourceSpreadsheet_();
  const sourceSheet = ss.getSheetByName(options.sourceSheetName);
  if (!sourceSheet) throw new Error('No se encontró la hoja "' + options.sourceSheetName + '".');

  const props = PropertiesService.getScriptProperties();
  const existingFileId = (props.getProperty(options.fileIdProp) || '').trim();

  let targetSs;
  let targetFileId = existingFileId;
  if (isReusableDownloadFile_(existingFileId, options.fileName)) {
    targetSs = SpreadsheetApp.openById(existingFileId);
  } else {
    targetSs = SpreadsheetApp.create(options.fileName);
    targetFileId = targetSs.getId();
  }

  const stagingName = '__NEW_' + options.outputSheetName.replace(/[^A-Za-z0-9_]/g, '_') +
                      '_' + Date.now() + '_' + Utilities.getUuid().slice(0, 8);

  const copied = sourceSheet.copyTo(targetSs);
  copied.setName(stagingName);
  copied.showSheet();
  cleanExportSheetForDownload_(copied, options.filterHeaderRow || 1);

  targetSs.getSheets().forEach(sh => {
    if (sh.getSheetId() !== copied.getSheetId()) targetSs.deleteSheet(sh);
  });

  copied.setName(options.outputSheetName);
  targetSs.setActiveSheet(copied);
  targetSs.moveActiveSheet(1);

  SpreadsheetApp.flush();
  Utilities.sleep(DOWNLOAD_EXPORT_SETTLE_MS);

  props.setProperty(options.fileIdProp, targetFileId);
  props.setProperty(options.sourceVersionProp, String(getSourceSpreadsheetVersion_(ss)));
  return targetFileId;
}

function refreshSingleSheetDownload_(options, lockAlreadyHeld) {
  let lock = null;
  if (!lockAlreadyHeld) {
    lock = LockService.getScriptLock();
    lock.waitLock(DOWNLOAD_LOCK_TIMEOUT_MS);
  }
  try {
    return refreshSingleSheetDownloadCore_(options);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function isDownloadCurrent_(options, ss) {
  const props = PropertiesService.getScriptProperties();
  const fileId = (props.getProperty(options.fileIdProp) || '').trim();
  const exported = Number(props.getProperty(options.sourceVersionProp) || '0');
  return fileId && exported >= getSourceSpreadsheetVersion_(ss) && isReusableDownloadFile_(fileId, options.fileName)
    ? fileId : '';
}

function ensureSingleSheetDownloadCurrent_(options) {
  const ss = getDownloadSourceSpreadsheet_();
  const current = isDownloadCurrent_(options, ss);
  if (current) return current;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(DOWNLOAD_LOCK_TIMEOUT_MS)) {
    throw new Error('La descarga se está actualizando. Espera unos segundos y vuelve a intentarlo.');
  }
  try {
    // Otra ejecución pudo haberla actualizado mientras esperábamos el lock.
    return isDownloadCurrent_(options, ss) || refreshSingleSheetDownloadCore_(options);
  } finally {
    lock.releaseLock();
  }
}

function ensureHorarioDownloadReady_() {
  return ensureSingleSheetDownloadCurrent_(horarioDownloadOptions_());
}

function refreshHorarioOnlyDownloadFile_(lockAlreadyHeld) {
  return refreshSingleSheetDownload_(horarioDownloadOptions_(), !!lockAlreadyHeld);
}

function resetHorarioOnlyDownloadFile() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(HORARIO_ONLY_DOWNLOAD_FILE_ID_PROP);
  props.deleteProperty(HORARIO_EXPORT_SOURCE_VERSION_PROP);
  const newId = refreshHorarioOnlyDownloadFile_();
  SpreadsheetApp.getUi().alert('Archivo de descarga HORARIO regenerado.\n\nFile ID:\n' + newId);
}

function exportSpreadsheetAsXlsxBlob_(spreadsheetId, fileName) {
  const id = String(spreadsheetId || '').trim();
  if (!id) throw new Error('No se pudo identificar el spreadsheet a exportar.');

  const response = UrlFetchApp.fetch(
    'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(id) + '/export?format=xlsx',
    {
      method: 'get',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
      followRedirects: true
    }
  );

  const code = response.getResponseCode();
  const bytes = response.getBlob().getBytes();
  if (code < 200 || code >= 300) {
    throw new Error('No se pudo exportar a Excel. HTTP ' + code + '. ' + response.getContentText().slice(0, 500));
  }
  if (!bytes || !bytes.length) throw new Error('La exportación a Excel devolvió un archivo vacío.');

  return Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName || 'archivo.xlsx');
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function buildInlineXlsxDownloadHtml_(opts) {
  const fileName = String(opts.fileName || 'archivo.xlsx');
  const message = String(opts.message || 'El archivo se descargará automáticamente en unos segundos');
  const xlsxBlob = exportSpreadsheetAsXlsxBlob_(opts.fileId, fileName);
  const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  // El XLSX se exporta con el token del propietario y se entrega como Blob
  // local en el navegador: no se requiere compartir el archivo en Drive.
  const html = `<!doctype html>
<html><head><base target="_top"><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Descarga Excel</title>
<style>
  body{font-family:Arial,sans-serif;margin:0;padding:24px;color:#222;line-height:1.45}
  .card{max-width:560px;margin:0 auto;border:1px solid #e0e0e0;border-radius:12px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,.06)}
  h2{margin-top:0;font-size:22px} .note{margin-top:16px;font-size:14px;color:#555}
  #status{margin-top:14px;font-size:13px;color:#555} #error{margin-top:14px;font-size:13px;color:#b00020;white-space:pre-wrap}
</style></head>
<body><div class="card">
  <h2>Archivo listo para descargar</h2>
  <p>${escapeHtml_(message)}</p>
  <p class="note">Si la descarga no comienza, haz click <a id="downloadLink" href="#">aquí</a>.</p>
  <p id="status"></p><p id="error"></p>
</div>
<script>
(function () {
  var b64 = ${JSON.stringify(Utilities.base64Encode(xlsxBlob.getBytes()))};
  var fileName = ${JSON.stringify(fileName)};
  var mime = ${JSON.stringify(mime)};
  var started = false;
  function toBlob() {
    var bin = atob(b64), arr = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return new Blob([arr], { type: mime });
  }
  function go() {
    started = true;
    try {
      var url = URL.createObjectURL(toBlob());
      var a = document.createElement('a');
      a.href = url; a.download = fileName; a.style.display = 'none';
      document.body.appendChild(a); a.click();
      setTimeout(function () { try { URL.revokeObjectURL(url); a.remove(); } catch (e) {} }, 10000);
      document.getElementById('status').textContent = 'Descarga iniciada.';
    } catch (err) {
      document.getElementById('error').textContent = 'No se pudo iniciar la descarga: ' + (err && err.message ? err.message : err);
    }
  }
  document.getElementById('downloadLink').addEventListener('click', function (e) { e.preventDefault(); go(); });
  var ua = navigator.userAgent || '';
  var isIOS = /iPad|iPhone|iPod/.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  var isSafari = /Safari/i.test(ua) && !/Chrome|Chromium|CriOS|FxiOS|EdgiOS|OPR|Opera|Android/i.test(ua);
  if (!isIOS && !isSafari) setTimeout(function () { if (!started) go(); }, 900);
})();
</script></body></html>`;

  return HtmlService.createHtmlOutput(html).setTitle('Descarga Excel');
}

function buildDownloadErrorHtml_(title, err) {
  const detail = err && err.message ? err.message : String(err || 'Error desconocido');
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<h3>' + escapeHtml_(title) + '</h3><pre style="white-space:pre-wrap">' + escapeHtml_(detail) + '</pre></body></html>'
  );
}

/*******************************
 * HELPERS USADOS POR Horario.gs
 *******************************/
function normalizeRut_(value) {
  return String(value == null ? '' : value).toUpperCase().replace(/[^0-9K]/g, '');
}

function isProgramarActive_(value) {
  if (value === null || value === undefined || value === '') return false;
  const txt = String(value).trim().toUpperCase();
  if (!txt) return false;
  if (['0', '0.0', '0,0', '-', 'NO', 'N/A', 'FALSE'].includes(txt)) return false;
  const numeric = Number(txt.replace(',', '.'));
  if (!isNaN(numeric)) return numeric > 0;
  return true;
}

function normalizeMeetingTypeFromMaestro_(tipoRaw) {
  const t = String(tipoRaw || '').trim().toUpperCase();
  if (!t) return '';
  if (t === 'LABT' || t === 'LAB' || t === 'TALLER' || t === 'LAB/TALLER') return 'LAB/TALLER';
  if (t === 'OLIN' || t === 'CLSS') return 'CLAS';
  if (t === 'AYON') return 'AYUD';
  return t;
}

function getWeeklyMeetingTypeSet_(row, mIdx, days) {
  const set = new Set();
  days.forEach(dayName => {
    const dayIndex = mIdx[dayName];
    if (dayIndex == null || dayIndex < 0) return;
    const cell = row[dayIndex];
    if (!cell) return;
    String(cell).split(',').forEach(partRaw => {
      const part = String(partRaw || '').trim();
      if (!part) return;
      const tipo = normalizeMeetingTypeFromMaestro_(part.split(/\s+/)[0]);
      if (tipo) set.add(tipo);
    });
  });
  return set;
}

function isOnlyLabtEffective_(row, mIdx, days, hasClasesProgramar, hasAyudProgramar, hasLabtProgramar) {
  const weeklyTypes = getWeeklyMeetingTypeSet_(row, mIdx, days);
  const hasActualLabt = weeklyTypes.has('LAB/TALLER');
  const hasActualClassOrAyud = weeklyTypes.has('CLAS') || weeklyTypes.has('AYUD');
  if (hasActualLabt || hasActualClassOrAyud) return hasActualLabt && !hasActualClassOrAyud;
  return !hasClasesProgramar && !hasAyudProgramar && hasLabtProgramar;
}

function normalizeDateOutput_(v, tz) {
  if (v == null || v === '') return '';
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz || Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const parsed = parseDateFlexible_(v, tz);
  return parsed ? Utilities.formatDate(parsed, tz || Session.getScriptTimeZone(), 'dd/MM/yyyy') : String(v).trim();
}

function normalizeHorarioDateColumns_(sheet) {
  if (!sheet) return;
  let block;
  try { block = getHorarioBlock_(sheet); } catch (e) { return; }

  const idxIni = block.idx['INICIO'];
  const idxFin = block.idx['FIN'];
  if (idxIni == null || idxFin == null) return;

  const rowCount = Math.max(0, block.values.length - 1);
  if (!rowCount) return;

  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const range = sheet.getRange(block.dataStartRow, 1, rowCount, block.lastCol);
  const values = range.getValues();
  values.forEach(r => {
    r[idxIni] = normalizeDateOutput_(r[idxIni], tz);
    r[idxFin] = normalizeDateOutput_(r[idxFin], tz);
  });
  range.setValues(values);
  sheet.getRange(block.dataStartRow, idxIni + 1, rowCount, 1).setNumberFormat('@');
  sheet.getRange(block.dataStartRow, idxFin + 1, rowCount, 1).setNumberFormat('@');
}

/*******************************
 * LOG
 *******************************/
function logWebPublish_(action, status, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('LOG_PUBLICACION_WEB');
  if (!sh) {
    sh = ss.insertSheet('LOG_PUBLICACION_WEB');
    sh.appendRow(['TIMESTAMP', 'USER', 'ACTION', 'STATUS', 'DETAIL']);
  }
  sh.appendRow([new Date(), Session.getActiveUser().getEmail(), action, status, detail || '']);
}

/*******************************
 * MENU
 *******************************/
function onOpen() {
  ensureParametrosHorarioSheet_();

  SpreadsheetApp.getUi()
    .createMenu('Macros')
    .addItem('Crear/Actualizar PROFESORES_ING', 'copyProfesores')
    .addItem('Ingresar Nombres Banner', 'updateBannerNames')
    .addItem('Segundo Profe y Labt', 'syncSecondAndLabt')
    .addSeparator()
    .addItem('Crear/Actualizar Parámetros Horario', 'ensureParametrosHorarioSheet_')
    .addItem('Crear Horario', 'createHorario')
    .addItem('Llenar Salas Horario', 'fillSalasHorario')
    .addItem('Actualizar Salas Pruebas', 'updateSalasPruebas')
    .addSeparator()
    .addItem('Importar inscripciones UG307', 'importarInscripcionesUG307')
    .addSeparator()
    .addItem('Configurar MAESTRO fuente', 'setupSourceMaestroId')
    .addItem('Ver MAESTRO fuente actual', 'showSourceMaestroId')
    .addSeparator()
    .addItem('Configurar URL Web App', 'setupWebAppUrl')
    .addItem('Regenerar archivo Excel HORARIO', 'resetHorarioOnlyDownloadFile')
    .addItem('Diagnosticar', 'diagnoseWebPublishState')
    .addItem('Publicar/Actualizar web', 'publishAllStaticFilesToGitHub')
    .addSeparator()
    .addItem('Configurar GitHub', 'setupGitHubPublishing')
    .addItem('Probar conexión GitHub', 'testGitHubConnection')
    .addItem('Ver configuración', 'showGitHubConfigSummary')
    .addToUi();
}
