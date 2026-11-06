/**
 * Horario.gs — lógica de construcción del horario.
 *
 * Copiado sin cambios del script anterior (desde normalizeName hasta
 * updateSalasPruebas). Depende de funciones y constantes definidas en Web.gs
 * (getSourceSpreadsheetId_, getHorarioParams_, normalizeRut_,
 * isProgramarActive_, isOnlyLabtEffective_, normalizeHorarioDateColumns_,
 * LOCAL_PROFESORES_ING_SHEET_NAME, SOURCE_PROFESORES_SHEET_NAME).
 */

function normalizeName(str) {
  return str
    ? str.toString()
         .normalize('NFD')
         .replace(/[̀-ͯ]/g, '')
         .toUpperCase()
    : '';
}

// --- Helpers para MATERIA / CURSO / CODIGO -----------------------------

function normalizeMateriaCode(val) {
  return val == null ? '' : String(val).trim().toUpperCase();
}

function normalizeCurso(val) {
  if (val == null) return '';
  let s = String(val).trim();

  if (s.startsWith("'")) s = s.slice(1);
  if (!s) return '';

  if (/^\d+$/.test(s)) return s.padStart(4, '0');

  return s.toUpperCase();
}

function splitCodigo(codigoVal, materiaVal, cursoVal) {
  let materia = normalizeMateriaCode(materiaVal);
  let curso   = normalizeCurso(cursoVal);
  let codigoStr = '';

  if (codigoVal != null && codigoVal !== '') {
    codigoStr = String(codigoVal).trim().toUpperCase();

    const m = codigoStr.match(/^([A-ZÁÉÍÓÚÑ]+)(\d+)$/i);
    if (m) {
      materia = normalizeMateriaCode(m[1]);
      curso   = normalizeCurso(m[2]);
    }
  }

  if (!codigoStr) codigoStr = materia + curso;

  return { codigo: codigoStr, materia: materia, curso: curso };
}

function normalizeHeaderKey_(value) {
  return String(value == null ? '' : value)
    .replace(/ /g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function buildHeaderIndexNormalized_(headers) {
  const out = {};
  headers.forEach((h, i) => {
    const key = normalizeHeaderKey_(h);
    if (key && !(key in out)) out[key] = i;
  });
  return out;
}

function getFirstExistingIndex_(idxNorm, aliases) {
  for (const alias of aliases) {
    const key = normalizeHeaderKey_(alias);
    if (key in idxNorm) return idxNorm[key];
  }
  return -1;
}

function valueByAliases_(row, idxNorm, aliases) {
  const i = getFirstExistingIndex_(idxNorm, aliases);
  return i >= 0 ? row[i] : '';
}

function maestroProfessorFieldIndexes_(headers) {
  const idxNorm = buildHeaderIndexNormalized_(headers);
  return {
    idxNorm,
    rut1: getFirstExistingIndex_(idxNorm, ['RUT PROFESOR 1']),
    nombre1: getFirstExistingIndex_(idxNorm, [
      'NOMBRE PROFESOR BANNER 1 (PROFESOR PRINCIPAL SESIÓN 01)',
      'NOMBRE PROFESOR BANNER 1 \n(PROFESOR PRINCIPAL SESIÓN 01)',
      'NOMBRE PROFESOR BANNER 1'
    ]),
    email1: getFirstExistingIndex_(idxNorm, ['EMAIL PROFESOR 1']),
    telefono1: getFirstExistingIndex_(idxNorm, ['TELEFONO PROFESOR 1']),

    rut2: getFirstExistingIndex_(idxNorm, ['RUT PROFESOR 2']),
    nombre2: getFirstExistingIndex_(idxNorm, [
      'NOMBRE PROFESOR 2 (2DO PROFESOR - SESIÓN 02)',
      'NOMBRE PROFESOR 2\n(2DO PROFESOR - SESIÓN 02)',
      'NOMBRE PROFESOR 2'
    ]),
    email2: getFirstExistingIndex_(idxNorm, ['EMAIL PROFESOR 2']),
    telefono2: getFirstExistingIndex_(idxNorm, ['TELEFONO PROFESOR 2']),

    rutLabt: getFirstExistingIndex_(idxNorm, ['RUT PROFESOR LABT']),
    nombreLabt: getFirstExistingIndex_(idxNorm, ['PROFESOR LABT', 'PROFESOR LABT ']),
    emailLabt: getFirstExistingIndex_(idxNorm, ['MAIL PROFESOR LABT']),

    sct: getFirstExistingIndex_(idxNorm, ['SCT-CHILE', 'SCTChile']),
    habilidades: getFirstExistingIndex_(idxNorm, ['Habilidades Transversales'])
  };
}

// ----------------------------------------------------------------------
// HELPERS HORARIO
// ----------------------------------------------------------------------

function findHorarioHeaderRow_(sheet) {
  const values = sheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i].map(v => String(v || '').trim());
    if (
      row.includes('AREA') &&
      row.includes('PLAN DE ESTUDIO') &&
      row.includes('NRC') &&
      row.includes('TIPO DE REUNION') &&
      row.includes('PROFESOR')
    ) {
      return i + 1; // 1-based
    }
  }
  return -1;
}

function getHorarioBlock_(sheet) {
  const headerRow = findHorarioHeaderRow_(sheet);
  if (headerRow < 1) {
    throw new Error('No se encontró la fila de encabezados de HORARIO.');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const numRows = lastRow - headerRow + 1;
  if (numRows < 1) {
    throw new Error('HORARIO no tiene datos bajo el encabezado.');
  }

  const values = sheet.getRange(headerRow, 1, numRows, lastCol).getValues();
  const headers = values[0];
  const idx = {};
  headers.forEach((h, i) => idx[String(h).trim()] = i);

  return {
    headerRow,
    values,
    headers,
    idx,
    dataRows: values.slice(1),
    dataStartRow: headerRow + 1,
    lastCol
  };
}

function setupHorarioLegend_(sheet) {
  sheet.clear();

  // Row 1
  sheet.getRange('B1:D1').merge();
  sheet.getRange('B1').setValue('Evaluaciones en las tardes o sábados o exámenes');
  sheet.getRange('A1').setBackground('#FCE4D6');

  // Row 2
  sheet.getRange('B2:D2').merge();
  sheet.getRange('B2').setValue('Evaluaciones en horario de Clase/Ayudantía');
  sheet.getRange('A2').setBackground('#CCEEFF');

  // Legend
  sheet.getRange('A4').setValue('TIPO HORARIO');
  sheet.getRange('B4').setValue('DESCRIPCION');

  const legendRows = [
    ['CLAS', 'CLASES'],
    ['OLIN', 'CLASES ONLINE'],
    ['AYUD', 'AYUDANTIA'],
    ['AYON', 'AYUDANTIA ONLINE'],
    ['LABT', 'LABORATORIO'],
    ['EXAM', 'EXAMEN'],
    ['PRBA', 'PRUEBA (1, 2, 3...)']
  ];
  sheet.getRange(5, 1, legendRows.length, 2).setValues(legendRows);

  sheet.getRange('B1').setFontWeight('bold');
  sheet.getRange('B2').setFontWeight('bold');
  sheet.getRange('A4:B4').setFontWeight('bold');

  sheet.setFrozenRows(0);
}

function parseDateFlexible_(v, tz) {
  if (v instanceof Date) return new Date(v.getTime());
  const s = String(v || '').trim();
  if (!s) return null;

  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function formatDateDDMMYYYY_(d, tz) {
  return Utilities.formatDate(d, tz || Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function sameDay_(a, b) {
  return a && b &&
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate();
}

function addDays_(d, n) {
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + n);
  return x;
}

function minDate_(a, b) {
  if (!a) return b;
  if (!b) return a;
  return a.getTime() <= b.getTime() ? a : b;
}

function maxDate_(a, b) {
  if (!a) return b;
  if (!b) return a;
  return a.getTime() >= b.getTime() ? a : b;
}

function intervalsOverlapOrTouch_(aStart, aEnd, bStart, bEnd) {
  return addDays_(aEnd, 1).getTime() >= bStart.getTime() &&
         addDays_(bEnd, 1).getTime() >= aStart.getTime();
}

function mergeIntervals_(intervals) {
  if (!intervals.length) return [];
  const sorted = intervals
    .map(x => ({
      start: new Date(x.start.getTime()),
      end: new Date(x.end.getTime()),
      sala: x.sala || ''
    }))
    .sort((a, b) => a.start - b.start || a.end - b.end);

  const out = [sorted[0]];
  for (let i = 1; i < sorted.length; i++) {
    const cur = sorted[i];
    const last = out[out.length - 1];

    if (last.sala === cur.sala && intervalsOverlapOrTouch_(last.start, last.end, cur.start, cur.end)) {
      last.end = maxDate_(last.end, cur.end);
    } else {
      out.push(cur);
    }
  }
  return out;
}

function getDayNameFromHorarioRow_(row, idx) {
  const days = ['Lunes','Martes','Miercoles','Jueves','Viernes'];
  for (const d of days) {
    const i = idx[d];
    if (i != null && String(row[i] || '').trim() !== '') return d;
  }
  return '';
}

function getTimeFromHorarioRow_(row, idx) {
  const day = getDayNameFromHorarioRow_(row, idx);
  if (!day) return '';
  return String(row[idx[day]] || '').trim();
}

function getTipoNormalized_(tipo) {
  let t = String(tipo || '').trim().toUpperCase();
  if (t === 'LABT') t = 'LAB/TALLER';
  return t;
}

function getDayColumnNrcPeriodo_(row, pIdx) {
  const days = ['LUNES','MARTES','MIERCOLES','JUEVES','VIERNES'];
  for (const d of days) {
    if (pIdx[d] != null) {
      const v = String(row[pIdx[d]] || '').trim();
      if (v) return { day: d, time: v };
    }
  }
  return { day: '', time: '' };
}

function parseDDMMYYYY_(s) {
  const m = String(s || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

function shiftDateString_(dateStr, days, tz) {
  const d = parseDDMMYYYY_(dateStr);
  if (!d) return dateStr;
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, tz || Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function previousWeekdayBeforeDate_(targetDate, weekdayIndex) {
  const d = new Date(targetDate.getTime());
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() - 1);

  while (d.getDay() !== weekdayIndex) {
    d.setDate(d.getDate() - 1);
  }

  return d;
}

function latestExamDateFromMaestro_(maestroRow, maestroHdr, mIdx, tz) {
  const obsI = mIdx['OBSERVACION'];
  let latest = null;

  for (let c = obsI + 1; c < maestroHdr.length; c++) {
    const headerCell = maestroHdr[c];
    if (!(headerCell instanceof Date)) continue;

    const cell = maestroRow[c];
    if (!cell) continue;

    const parts = String(cell).split(',');
    for (const partRaw of parts) {
      const part = String(partRaw || '').trim().toUpperCase();
      if (!part) continue;
      const tipo = part.split(' ')[0];
      if (tipo === 'EXAM') {
        const examDate = new Date(headerCell.getTime());
        latest = maxDate_(latest, examDate);
      }
    }
  }
  return latest;
}

function buildFacultadCegExamMap_(mSheet, tz) {
  const all = mSheet.getDataRange().getValues();
  if (all.length < 2) return {};

  const hdr = all[0];
  const mIdx = hdr.reduce((m,h,i)=>(m[h]=i,m), {});
  const out = {};

  if (!('NRC' in mIdx) || !('AREA' in mIdx) || !('OBSERVACION' in mIdx)) {
    return out;
  }

  for (let r = 1; r < all.length; r++) {
    const row = all[r];
    const nrc = String(row[mIdx['NRC']] || '').trim();
    const area = String(row[mIdx['AREA']] || '').trim().toUpperCase();
    if (!nrc || area !== 'FACULTAD CEG') continue;

    const latestExam = latestExamDateFromMaestro_(row, hdr, mIdx, tz);
    if (!latestExam) continue;

    out[nrc] = latestExam;
  }
  return out;
}

function maestroAreaMap_(mSheet) {
  const all = mSheet.getDataRange().getValues();
  if (all.length < 2) return {};
  const hdr = all[0];
  const mIdx = hdr.reduce((m,h,i)=>(m[h]=i,m), {});
  const out = {};
  if (!('NRC' in mIdx) || !('AREA' in mIdx)) return out;

  for (let r = 1; r < all.length; r++) {
    const nrc = String(all[r][mIdx['NRC']] || '').trim();
    const area = String(all[r][mIdx['AREA']] || '').trim();
    if (nrc) out[nrc] = area;
  }
  return out;
}

// ----------------------------------------------------------------------
// COPY PROFESORES
// ----------------------------------------------------------------------

function copyProfesores() {
  const srcSS  = SpreadsheetApp.openById(getSourceSpreadsheetId_());
  const mSheet = srcSS.getSheetByName('MAESTRO');
  const pSrc   = srcSS.getSheetByName(SOURCE_PROFESORES_SHEET_NAME);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let outSheet = ss.getSheetByName(LOCAL_PROFESORES_ING_SHEET_NAME);

  if (!outSheet) {
    outSheet = ss.insertSheet(LOCAL_PROFESORES_ING_SHEET_NAME);
  } else {
    outSheet.clearContents();
  }

  if (!mSheet) {
    SpreadsheetApp.getUi().alert('No encontré la hoja MAESTRO en el documento fuente.');
    return;
  }

  const data = mSheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('La hoja MAESTRO debe tener datos.');
    return;
  }

  const hIdx = data[0].reduce((m,h,i)=>(m[h]=i,m), {});
  const hIdxNorm = buildHeaderIndexNormalized_(data[0]);
  const profFields = maestroProfessorFieldIndexes_(data[0]);

  function getMaestroIndex_(aliases) {
    return getFirstExistingIndex_(hIdxNorm, aliases);
  }

  const idxMandante = getMaestroIndex_(['CURSO MANDANTE']);
  const idxSecc     = getMaestroIndex_(['SECCIONES', 'SECC.']);
  const idxLc       = getMaestroIndex_(['LC', 'LISTA CRUZADA']);
  const idxConn     = getMaestroIndex_(['CONECTOR DE LIGA', 'CONECTOR LIGA', 'CONECTOR  LIGA']);

  function cell_(row, idx) {
    return idx >= 0 ? row[idx] : '';
  }

  function text_(v) {
    return String(v == null ? '' : v).trim();
  }

  function hasValue_(v) {
    return text_(v) !== '';
  }

  function normalizeLcForCopy_(value) {
    return String(value == null ? '' : value).trim().toUpperCase();
  }

  function normalizeSectionForCopy_(value) {
    const s = String(value == null ? '' : value).trim();
    if (!s) return '';
    if (/^\d+(\.0+)?$/.test(s)) return String(Number(s));
    return s.toUpperCase();
  }

  function linkedKeyForCopy_(row) {
    const lc = normalizeLcForCopy_(cell_(row, idxLc));
    if (lc) return lc;
    return normalizeLcForCopy_(cell_(row, idxConn));
  }

  function professorCompleteness_(row) {
    let score = 0;
    if (hasValue_(cell_(row, profFields.rut1))) score += 16;
    if (hasValue_(cell_(row, profFields.nombre1))) score += 16;
    if (hasValue_(cell_(row, profFields.email1))) score += 4;
    if (hasValue_(cell_(row, profFields.telefono1))) score += 2;
    if (hasValue_(cell_(row, profFields.rut2))) score += 4;
    if (hasValue_(cell_(row, profFields.nombre2))) score += 4;
    if (hasValue_(cell_(row, profFields.rutLabt))) score += 4;
    if (hasValue_(cell_(row, profFields.nombreLabt))) score += 4;
    return score;
  }

  function hasAnyProfessorData_(row) {
    return professorCompleteness_(row) > 0;
  }

  function sourcePriority_(row) {
    const mandante = idxMandante >= 0 ? String(row[idxMandante] || '').trim().toUpperCase() : '';
    // Lower is better. Prefer CURSO MANDANTE = SI, then rows with more complete professor data.
    let priority = 1000;
    if (mandante === 'SI') priority -= 300;
    priority -= professorCompleteness_(row);
    return priority;
  }

  function putBestSource_(map, key, row) {
    if (!key || !hasAnyProfessorData_(row)) return;
    const rec = {
      row,
      priority: sourcePriority_(row)
    };
    const prev = map[key];
    if (!prev || rec.priority < prev.priority) map[key] = rec;
  }

  function firstNonBlank_(primary, fallback) {
    return hasValue_(primary) ? primary : fallback;
  }

  // TIPO DE PROFESOR continues to be based only on the external MAESTRO > PROFESORES sheet:
  // if the RUT appears in PROFESORES.RUT PROFESOR 1 => JORNADA; otherwise => HONORARIO.
  let jornada = new Set();
  if (pSrc) {
    const pData = pSrc.getDataRange().getValues();
    if (pData.length >= 2) {
      const pIdxNorm = buildHeaderIndexNormalized_(pData[0]);
      const rutCol = getFirstExistingIndex_(pIdxNorm, ['RUT PROFESOR 1']);
      if (rutCol >= 0) {
        jornada = new Set(
          pData.slice(1)
            .map(r => normalizeRut_(r[rutCol]))
            .filter(Boolean)
        );
      }
    }
  }

  // LC-aware source maps. These are used only to complete professor fields in rows
  // where MAESTRO leaves the linked / non-mandante row blank.
  const dataRows = data.slice(1);
  const byLcSection = {}; // LC|SECC
  const byLcOnly    = {}; // LC

  dataRows.forEach(row => {
    const lcKey = linkedKeyForCopy_(row);
    if (!lcKey) return;

    const seccKey = normalizeSectionForCopy_(cell_(row, idxSecc));
    putBestSource_(byLcOnly, lcKey, row);
    if (seccKey) putBestSource_(byLcSection, [lcKey, seccKey].join('|'), row);
  });

  function fallbackProfessorRowFor_(row) {
    if (!linkedKeyForCopy_(row)) return null;

    const lcKey = linkedKeyForCopy_(row);
    const seccKey = normalizeSectionForCopy_(cell_(row, idxSecc));

    if (seccKey && byLcSection[[lcKey, seccKey].join('|')]) {
      return byLcSection[[lcKey, seccKey].join('|')].row;
    }
    if (byLcOnly[lcKey]) return byLcOnly[lcKey].row;
    return null;
  }

  const HEADER = [
    'PREGRADO/POSTGRADO','CURSO MANDANTE','AREA','PLAN DE ESTUDIO','NRC',
    'CONECTOR DE LIGA','LC','CODIGO','MATERIA','CURSO','TITULO','SECCIONES',
    'SCT-CHILE','Habilidades Transversales','RUT PROFESOR 1',
    'NOMBRE PROFESOR BANNER 1 (PROFESOR PRINCIPAL SESIÓN 01)',
    'EMAIL PROFESOR 1','TIPO DE PROFESOR','TELEFONO PROFESOR 1',
    'RUT PROFESOR 2','NOMBRE PROFESOR 2 (2DO PROFESOR - SESIÓN 02)',
    'EMAIL PROFESOR 2','TELEFONO PROFESOR 2',
    'RUT PROFESOR LABT','PROFESOR LABT','MAIL PROFESOR LABT'
  ];
  outSheet.clearContents();
  outSheet.appendRow(HEADER);

  let inheritedLcCount = 0;

  const rows = dataRows.map(r => {
    const nivel = (r[hIdx['ELECTIVO']]||'').toString().toUpperCase() === 'POSTGRADO'
      ? 'POSTGRADO'
      : 'PREGRADO';

    const fallbackRow = fallbackProfessorRowFor_(r);
    const shouldUseFallback = fallbackRow && professorCompleteness_(fallbackRow) > professorCompleteness_(r);
    const pRow = shouldUseFallback ? fallbackRow : null;
    if (pRow) inheritedLcCount++;

    const rut1 = firstNonBlank_(cell_(r, profFields.rut1), pRow ? cell_(pRow, profFields.rut1) : '');
    const nombre1 = firstNonBlank_(cell_(r, profFields.nombre1), pRow ? cell_(pRow, profFields.nombre1) : '');
    const email1 = firstNonBlank_(cell_(r, profFields.email1), pRow ? cell_(pRow, profFields.email1) : '');
    const tel1 = firstNonBlank_(cell_(r, profFields.telefono1), pRow ? cell_(pRow, profFields.telefono1) : '');

    const rut2 = firstNonBlank_(cell_(r, profFields.rut2), pRow ? cell_(pRow, profFields.rut2) : '');
    const nombre2 = firstNonBlank_(cell_(r, profFields.nombre2), pRow ? cell_(pRow, profFields.nombre2) : '');
    const email2 = firstNonBlank_(cell_(r, profFields.email2), pRow ? cell_(pRow, profFields.email2) : '');
    const tel2 = firstNonBlank_(cell_(r, profFields.telefono2), pRow ? cell_(pRow, profFields.telefono2) : '');

    const rutLabt = firstNonBlank_(cell_(r, profFields.rutLabt), pRow ? cell_(pRow, profFields.rutLabt) : '');
    const nombreLabt = firstNonBlank_(cell_(r, profFields.nombreLabt), pRow ? cell_(pRow, profFields.nombreLabt) : '');
    const emailLabt = firstNonBlank_(cell_(r, profFields.emailLabt), pRow ? cell_(pRow, profFields.emailLabt) : '');

    const tipo = jornada.has(normalizeRut_(rut1)) ? 'JORNADA' : 'HONORARIO';

    const split = splitCodigo(r[hIdx['CODIGO']], r[hIdx['MATERIA']], r[hIdx['CURSO']]);
    const codigo  = split.codigo;
    const materia = split.materia;
    const curso   = split.curso;
    const cursoOut = curso ? "'" + curso : '';

    return [
      nivel,
      r[hIdx['CURSO MANDANTE']],
      r[hIdx['AREA']],
      r[hIdx['PLAN DE ESTUDIO']],
      r[hIdx['NRC']],
      r[hIdx['CONECTOR DE LIGA']],
      r[hIdx['LC']],
      codigo,
      materia,
      cursoOut,
      normalizeName(r[hIdx['TITULO']]),
      r[hIdx['SECCIONES']],
      profFields.sct >= 0 ? r[profFields.sct] : '',
      profFields.habilidades >= 0 ? r[profFields.habilidades] : '',
      rut1,
      normalizeName(nombre1),
      email1,
      tipo,
      tel1,
      rut2,
      normalizeName(nombre2),
      email2,
      tel2,
      rutLabt,
      normalizeName(nombreLabt),
      emailLabt
    ];
  });

  if (rows.length) {
    outSheet.getRange(2,1,rows.length,HEADER.length).setValues(rows);
    outSheet.getRange(2,10,rows.length,1).setNumberFormat('@');
  }

  SpreadsheetApp.getUi().alert(
    `✅ ${LOCAL_PROFESORES_ING_SHEET_NAME} generado.\n\n` +
    `Filas generadas: ${rows.length}\n` +
    `Filas completadas con lógica LC: ${inheritedLcCount}`
  );
}

// ----------------------------------------------------------------------
// UPDATE BANNER NAMES
// ----------------------------------------------------------------------

function updateBannerNames() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('NRC POR PERIODO');
  const tSheet = ss.getSheetByName(LOCAL_PROFESORES_ING_SHEET_NAME);

  if (!tSheet) {
    ui.alert(`Falta la hoja “${LOCAL_PROFESORES_ING_SHEET_NAME}”.`);
    return;
  }

  if (!pSheet) {
    ui.alert('No se encontró la hoja “NRC POR PERIODO”. Se mantendrán los nombres de profesores provenientes de MAESTRO en PROFESORES_ING.');
    return;
  }

  function normalizeNrcForKey_(value) {
    const s = String(value == null ? '' : value).trim();
    if (!s) return '';
    if (/^\d+(\.0+)?$/.test(s)) return String(Number(s));
    return s.toUpperCase();
  }

  function normalizeSectionForKey_(value) {
    const s = String(value == null ? '' : value).trim();
    if (!s) return '';
    if (/^\d+(\.0+)?$/.test(s)) return String(Number(s));
    return s.toUpperCase();
  }

  function normalizeLcForKey_(value) {
    return String(value == null ? '' : value).trim().toUpperCase();
  }

  function profPriority_(tipo) {
    const t = String(tipo || '').trim().toUpperCase();
    if (t === 'CLAS') return 1;
    if (t === 'AYUD') return 2;
    if (t === 'LBEX' || t === 'LABT' || t === 'LAB/TALLER') return 3;
    if (t === 'EXAM') return 4;
    if (t === 'PRBA' || t === 'TARDE') return 5;
    return 9;
  }

  function putBest_(map, key, rec) {
    if (!key) return;
    if (!rec.prof && !rec.rut) return;
    const prev = map[key];
    if (!prev || rec.priority < prev.priority) map[key] = rec;
  }

  function setIfDifferent_(sheet, rowNumber, colIndex0, newValue) {
    if (colIndex0 == null || colIndex0 < 0) return false;
    const next = String(newValue == null ? '' : newValue).trim();
    if (!next) return false;
    const range = sheet.getRange(rowNumber, colIndex0 + 1);
    const current = String(range.getValue() == null ? '' : range.getValue()).trim();
    if (current !== next) {
      range.setValue(next);
      return true;
    }
    return false;
  }

  const pd = pSheet.getDataRange().getValues();
  if (pd.length < 2) {
    ui.alert('La hoja “NRC POR PERIODO” no tiene datos para actualizar nombres.');
    return;
  }

  const pIdxNorm = buildHeaderIndexNormalized_(pd[0]);
  const pNrcI    = getFirstExistingIndex_(pIdxNorm, ['NRC']);
  const pMatI    = getFirstExistingIndex_(pIdxNorm, ['MATERIA']);
  const pCurI    = getFirstExistingIndex_(pIdxNorm, ['CURSO']);
  const pSecI    = getFirstExistingIndex_(pIdxNorm, ['SECC.', 'SECCIONES']);
  const pRutI    = getFirstExistingIndex_(pIdxNorm, ['RUT PROFESOR']);
  const pProfI   = getFirstExistingIndex_(pIdxNorm, ['PROFESOR']);
  const pTipoI   = getFirstExistingIndex_(pIdxNorm, ['TIPO', 'TIPO DE REUNION', 'TIPO DE REUNIÓN']);
  const pLcI     = getFirstExistingIndex_(pIdxNorm, ['LISTA CRUZADA', 'LC']);
  const pConnI   = getFirstExistingIndex_(pIdxNorm, ['CONECTOR LIGA', 'CONECTOR  LIGA', 'CONECTOR DE LIGA']);

  const requiredPeriodo = [
    ['NRC', pNrcI], ['MATERIA', pMatI], ['CURSO', pCurI], ['SECC.', pSecI],
    ['RUT PROFESOR', pRutI], ['PROFESOR', pProfI]
  ];
  for (const [name, idx] of requiredPeriodo) {
    if (idx < 0) {
      ui.alert(`Falta la columna "${name}" en NRC POR PERIODO.`);
      return;
    }
  }

  // Maps from NRC POR PERIODO.
  // The direct map is kept, but LC-aware maps are also built because many linked rows
  // in PROFESORES_ING either have blank RUTs or use "sin NRC asociado".
  const byExact       = {}; // NRC|MATERIA|CURSO|SECC|RUT
  const byCourse      = {}; // NRC|MATERIA|CURSO|SECC
  const byLcCourse    = {}; // LC|MATERIA|CURSO|SECC
  const byLcSection   = {}; // LC|SECC
  const byLcSectionRut= {}; // LC|SECC|RUT

  pd.slice(1).forEach(row => {
    const nrc  = normalizeNrcForKey_(row[pNrcI]);
    const mat  = normalizeMateriaCode(row[pMatI]);
    const cur  = normalizeCurso(row[pCurI]);
    const secc = normalizeSectionForKey_(row[pSecI]);
    const rut  = normalizeRut_(row[pRutI]);
    const prof = row[pProfI];
    const profStr = prof == null ? '' : String(prof).trim();
    const tipo = pTipoI >= 0 ? row[pTipoI] : '';
    const lc = normalizeLcForKey_(pLcI >= 0 ? row[pLcI] : '');
    const conn = normalizeLcForKey_(pConnI >= 0 ? row[pConnI] : '');
    const lcKeyValue = lc || conn;

    if (!profStr && !rut) return;

    const rec = {
      rut,
      prof: profStr,
      priority: profPriority_(tipo)
    };

    putBest_(byExact, [nrc, mat, cur, secc, rut].join('|'), rec);
    putBest_(byCourse, [nrc, mat, cur, secc].join('|'), rec);

    if (lcKeyValue) {
      putBest_(byLcCourse, [lcKeyValue, mat, cur, secc].join('|'), rec);
      putBest_(byLcSection, [lcKeyValue, secc].join('|'), rec);
      if (rut) putBest_(byLcSectionRut, [lcKeyValue, secc, rut].join('|'), rec);
    }
  });

  const fd = tSheet.getDataRange().getValues();
  if (fd.length < 2) {
    ui.alert(`La hoja “${LOCAL_PROFESORES_ING_SHEET_NAME}” no tiene datos.`);
    return;
  }

  const fIdxNorm = buildHeaderIndexNormalized_(fd[0]);
  const fIdx = {};
  fd[0].forEach((h, i) => { fIdx[String(h).trim()] = i; });

  const nrcI       = getFirstExistingIndex_(fIdxNorm, ['NRC']);
  const matI       = getFirstExistingIndex_(fIdxNorm, ['MATERIA']);
  const curI       = getFirstExistingIndex_(fIdxNorm, ['CURSO']);
  const seccI      = getFirstExistingIndex_(fIdxNorm, ['SECCIONES', 'SECC.']);
  const lcI        = getFirstExistingIndex_(fIdxNorm, ['LC', 'LISTA CRUZADA']);
  const connI      = getFirstExistingIndex_(fIdxNorm, ['CONECTOR DE LIGA', 'CONECTOR LIGA', 'CONECTOR  LIGA']);
  const mandanteI  = getFirstExistingIndex_(fIdxNorm, ['CURSO MANDANTE']);
  const rutI       = getFirstExistingIndex_(fIdxNorm, ['RUT PROFESOR 1']);
  const nameI      = getFirstExistingIndex_(fIdxNorm, ['NOMBRE PROFESOR BANNER 1 (PROFESOR PRINCIPAL SESIÓN 01)', 'NOMBRE PROFESOR BANNER 1']);
  const emailI     = getFirstExistingIndex_(fIdxNorm, ['EMAIL PROFESOR 1']);
  const tipoProfI  = getFirstExistingIndex_(fIdxNorm, ['TIPO DE PROFESOR']);
  const phoneI     = getFirstExistingIndex_(fIdxNorm, ['TELEFONO PROFESOR 1']);

  const requiredProfIng = [
    ['NRC', nrcI], ['MATERIA', matI], ['CURSO', curI], ['SECCIONES', seccI],
    ['RUT PROFESOR 1', rutI], ['NOMBRE PROFESOR BANNER 1 (PROFESOR PRINCIPAL SESIÓN 01)', nameI]
  ];
  for (const [name, idx] of requiredProfIng) {
    if (idx < 0) {
      ui.alert(`Falta la columna "${name}" en ${LOCAL_PROFESORES_ING_SHEET_NAME}.`);
      return;
    }
  }

  tSheet.getRange(2, 1, fd.length - 1, fd[0].length).setBackground(null);

  // Local LC map from PROFESORES_ING itself. This lets rows tagged CURSO MANDANTE = NO
  // inherit the RUT/email/tipo/teléfono from the linked SI row when MAESTRO left them blank.
  const localLcMap = {};       // LC|SECC
  const localLcRutMap = {};    // LC|SECC|RUT

  function localPriority_(row) {
    const mandante = mandanteI >= 0 ? String(row[mandanteI] || '').trim().toUpperCase() : '';
    const hasName = String(row[nameI] || '').trim() !== '';
    const hasRut = normalizeRut_(row[rutI]) !== '';
    let score = 50;
    if (mandante === 'SI') score -= 20;
    if (hasName) score -= 10;
    if (hasRut) score -= 5;
    return score;
  }

  function putBestLocal_(map, key, rec) {
    if (!key) return;
    if (!rec.name && !rec.rut) return;
    const prev = map[key];
    if (!prev || rec.priority < prev.priority) map[key] = rec;
  }

  fd.slice(1).forEach((row, i) => {
    const lc = normalizeLcForKey_((lcI >= 0 ? row[lcI] : '') || (connI >= 0 ? row[connI] : ''));
    const secc = normalizeSectionForKey_(row[seccI]);
    const rut = normalizeRut_(row[rutI]);
    const name = String(row[nameI] || '').trim();
    if (!lc || (!rut && !name)) return;

    const rec = {
      rut,
      name,
      email: emailI >= 0 ? String(row[emailI] || '').trim() : '',
      tipoProfesor: tipoProfI >= 0 ? String(row[tipoProfI] || '').trim() : '',
      phone: phoneI >= 0 ? String(row[phoneI] || '').trim() : '',
      priority: localPriority_(row),
      sourceRowNumber: i + 2
    };

    putBestLocal_(localLcMap, [lc, secc].join('|'), rec);
    if (rut) putBestLocal_(localLcRutMap, [lc, secc, rut].join('|'), rec);
  });

  let updated = 0;
  let lcFilled = 0;
  let directFilled = 0;
  let localCopied = 0;
  let missing = 0;

  fd.slice(1).forEach((row, r) => {
    const targetRow = r + 2;

    const nrc  = normalizeNrcForKey_(row[nrcI]);
    const mat  = normalizeMateriaCode(row[matI]);
    const cur  = normalizeCurso(row[curI]);
    const secc = normalizeSectionForKey_(row[seccI]);
    const rut  = normalizeRut_(row[rutI]);
    const lc   = normalizeLcForKey_((lcI >= 0 ? row[lcI] : '') || (connI >= 0 ? row[connI] : ''));

    const currentName = String(row[nameI] || '').trim();

    const directRec =
      (rut ? byExact[[nrc, mat, cur, secc, rut].join('|')] : null) ||
      byCourse[[nrc, mat, cur, secc].join('|')];

    const lcRec = lc ? (
      (rut ? byLcSectionRut[[lc, secc, rut].join('|')] : null) ||
      byLcCourse[[lc, mat, cur, secc].join('|')] ||
      byLcSection[[lc, secc].join('|')]
    ) : null;

    const localRec = lc ? (
      (rut ? localLcRutMap[[lc, secc, rut].join('|')] : null) ||
      localLcMap[[lc, secc].join('|')]
    ) : null;

    const chosen = directRec || lcRec;

    let changedThisRow = false;

    if (chosen) {
      if (!rut && chosen.rut) {
        changedThisRow = setIfDifferent_(tSheet, targetRow, rutI, chosen.rut) || changedThisRow;
      }

      if (chosen.prof) {
        changedThisRow = setIfDifferent_(tSheet, targetRow, nameI, chosen.prof) || changedThisRow;
      }

      if (directRec) directFilled++;
      else lcFilled++;
    }

    // If the Banner/NRC data did not provide every local administrative field,
    // use the linked LC row inside PROFESORES_ING as a complement.
    if (localRec && localRec.sourceRowNumber !== targetRow) {
      if (!String(row[rutI] || '').trim() && localRec.rut) {
        changedThisRow = setIfDifferent_(tSheet, targetRow, rutI, localRec.rut) || changedThisRow;
      }
      if (!String(row[emailI] || '').trim() && localRec.email) {
        changedThisRow = setIfDifferent_(tSheet, targetRow, emailI, localRec.email) || changedThisRow;
      }
      if (tipoProfI >= 0) {
        const currentTipo = String(row[tipoProfI] || '').trim().toUpperCase();
        if ((!currentTipo || currentTipo === 'HONORARIO') && localRec.tipoProfesor) {
          changedThisRow = setIfDifferent_(tSheet, targetRow, tipoProfI, localRec.tipoProfesor) || changedThisRow;
        }
      }
      if (phoneI >= 0 && !String(row[phoneI] || '').trim() && localRec.phone) {
        changedThisRow = setIfDifferent_(tSheet, targetRow, phoneI, localRec.phone) || changedThisRow;
      }

      if (!chosen && localRec.name && !currentName) {
        changedThisRow = setIfDifferent_(tSheet, targetRow, nameI, localRec.name) || changedThisRow;
      }

      if (changedThisRow) localCopied++;
    }

    if (changedThisRow) {
      updated++;
      return;
    }

    const finalName = String(tSheet.getRange(targetRow, nameI + 1).getValue() || '').trim();
    if (!finalName) {
      tSheet.getRange(targetRow, nameI + 1).setBackground('#ffff00');
      missing++;
    }
  });

  SpreadsheetApp.flush();

  ui.alert(
    `✅ Filas actualizadas: ${updated}\n` +
    `• Match directo NRC/curso: ${directFilled}\n` +
    `• Match por LC/lista cruzada: ${lcFilled}\n` +
    `• Complementos copiados desde fila LC en PROFESORES_ING: ${localCopied}\n` +
    `⚠️ Sin nombre encontrado: ${missing}`
  );
}

// ----------------------------------------------------------------------
// SYNC SECOND AND LABT
// ----------------------------------------------------------------------

function syncSecondAndLabt() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOCAL_PROFESORES_ING_SHEET_NAME);
  if (!sheet){ ui.alert(`No encontré "${LOCAL_PROFESORES_ING_SHEET_NAME}".`); return; }

  const data = sheet.getDataRange().getValues();
  if (data.length<2){ ui.alert('La hoja no tiene datos.'); return; }

  const idx = data[0].reduce((m,h,i)=>(m[h]=i,m), {});
  const need= ['RUT PROFESOR 1',
               'NOMBRE PROFESOR BANNER 1 (PROFESOR PRINCIPAL SESIÓN 01)',
               'RUT PROFESOR 2','NOMBRE PROFESOR 2 (2DO PROFESOR - SESIÓN 02)',
               'RUT PROFESOR LABT','PROFESOR LABT'];
  for (let c of need) if (!(c in idx)){
    ui.alert(`Falta columna "${c}".`); return;
  }

  const bannerMap = {};
  data.slice(1).forEach(r=>{
    const rut = r[idx['RUT PROFESOR 1']];
    const bn  = r[idx['NOMBRE PROFESOR BANNER 1 (PROFESOR PRINCIPAL SESIÓN 01)']];
    if (rut&&bn) bannerMap[rut]=bn;
  });

  const nRows = data.length-1;
  sheet.getRange(2,idx['NOMBRE PROFESOR 2 (2DO PROFESOR - SESIÓN 02)']+1,nRows,1)
       .setBackground(null);
  sheet.getRange(2,idx['PROFESOR LABT']+1,nRows,1)
       .setBackground(null);

  let upd2=0,mis2=0,updL=0,misL=0;
  data.slice(1).forEach((r,i)=>{
    const rNum = i+2;
    const rut2 = r[idx['RUT PROFESOR 2']];
    const labt = r[idx['RUT PROFESOR LABT']];

    if (rut2) {
      if (bannerMap[rut2]) {
        sheet.getRange(rNum, idx['NOMBRE PROFESOR 2 (2DO PROFESOR - SESIÓN 02)']+1)
             .setValue(bannerMap[rut2]);
        upd2++;
      } else {
        sheet.getRange(rNum, idx['NOMBRE PROFESOR 2 (2DO PROFESOR - SESIÓN 02)']+1)
             .setBackground('#FFCCCC');
        mis2++;
      }
    }

    if (labt) {
      if (bannerMap[labt]) {
        sheet.getRange(rNum, idx['PROFESOR LABT']+1)
             .setValue(bannerMap[labt]);
        updL++;
      } else {
        sheet.getRange(rNum, idx['PROFESOR LABT']+1)
             .setBackground('#FFCCCC');
        misL++;
      }
    }
  });

  SpreadsheetApp.flush();
  ui.alert(`Segundo Profe: ✓${upd2} ✗${mis2}\nLabT: ✓${updL} ✗${misL}`);
}


// ----------------------------------------------------------------------
// HELPERS CURSO MANDANTE / LISTA CRUZADA
// ----------------------------------------------------------------------

function getLinkedCourseKey_(row, mIdx) {
  const lcIdx = mIdx['LC'];
  const conectorIdx = mIdx['CONECTOR DE LIGA'];

  const lc = lcIdx != null ? String(row[lcIdx] || '').trim().toUpperCase() : '';
  if (lc) return lc;

  const conector = conectorIdx != null ? String(row[conectorIdx] || '').trim().toUpperCase() : '';
  return conector;
}

function rowHasHorarioData_(row, mIdx, hdr, obsI) {
  const days = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes'];

  for (const d of days) {
    const idx = mIdx[d];
    if (idx != null && String(row[idx] || '').trim() !== '') return true;
  }

  if (obsI != null && obsI >= 0) {
    for (let c = obsI + 1; c < hdr.length; c++) {
      const headerCell = hdr[c];
      if (!(headerCell instanceof Date)) continue;
      if (String(row[c] || '').trim() !== '') return true;
    }
  }

  return false;
}

function buildMandanteHorarioSourceMap_(dataRows, mIdx, hdr, obsI) {
  const out = {};

  dataRows.forEach(row => {
    const mandante = String(row[mIdx['CURSO MANDANTE']] || '').trim().toUpperCase();
    if (mandante !== 'SI') return;

    const key = getLinkedCourseKey_(row, mIdx);
    if (!key) return;

    if (!rowHasHorarioData_(row, mIdx, hdr, obsI)) return;

    // There should normally be one SI row per LC/connector/section.
    // If there are duplicates, keep the first one to avoid unstable behavior.
    if (!out[key]) out[key] = row;
  });

  return out;
}

function getHorarioSourceRow_(row, mandanteHorarioSourceMap, mIdx, hdr, obsI) {
  if (rowHasHorarioData_(row, mIdx, hdr, obsI)) return row;

  const key = getLinkedCourseKey_(row, mIdx);
  if (key && mandanteHorarioSourceMap[key]) {
    return mandanteHorarioSourceMap[key];
  }

  return row;
}


// ----------------------------------------------------------------------
// CREATE HORARIO
// ----------------------------------------------------------------------

function createHorario() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = SpreadsheetApp.openById(getSourceSpreadsheetId_());
  const mSheet = src.getSheetByName('MAESTRO');
  if (!mSheet) {
    ui.alert('No encontré la hoja MAESTRO en el documento fuente.');
    return;
  }

  const params = getHorarioParams_();
  const classStart = params.classStart;
  const classEnd = params.classEnd;
  const ayudStart = params.ayudStart;
  const ayudEnd = params.ayudEnd;

  if (!classStart || !classEnd || !ayudStart || !ayudEnd) {
    throw new Error(
      'Faltan parámetros en la hoja PARAMETROS_HORARIO. ' +
      'Ejecuta “Crear/Actualizar Parámetros Horario” y completa los valores.'
    );
  }

  let horario = ss.getSheetByName('HORARIO');
  if (horario) {
    horario.clear();
  } else {
    horario = ss.insertSheet('HORARIO');
  }

  setupHorarioLegend_(horario);

  const HEADER = [
    'AREA','PLAN DE ESTUDIO','NRC','CONECTOR DE LIGA','LC','MATERIA','CURSO','SECC.','TITULO',
    'Lunes','Martes','Miercoles','Jueves','Viernes',
    'INICIO','FIN','SALA','TIPO DE REUNION','PROFESOR'
  ];

  const tableHeaderRow = 15;
  horario.getRange(tableHeaderRow, 1, 1, HEADER.length).setValues([HEADER]);

  const all = mSheet.getDataRange().getValues();
  const hdr = all[0];
  const mIdx = hdr.reduce((m,h,i) => (m[h] = i, m), {});
  const profFields = maestroProfessorFieldIndexes_(hdr);
  const obsI = mIdx['OBSERVACION'];
  const days = ['Lunes','Martes','Miercoles','Jueves','Viernes'];
  const tz   = ss.getSpreadsheetTimeZone();

  const requiredProgramarCols = [
    'Clases A PROGRAMAR',
    'Ayudantías PROGRAMAR',
    'Laboratorios o Talleres PROGRAMAR'
  ];
  for (const col of requiredProgramarCols) {
    if (!(col in mIdx)) {
      ui.alert(`No encontré la columna "${col}" en MAESTRO.`);
      return;
    }
  }

  if (!('Plan Común' in mIdx)) {
    ui.alert('No encontré la columna "Plan Común" en MAESTRO.');
    return;
  }

  const facultadCegExamMap = buildFacultadCegExamMap_(mSheet, tz);
  const dataRows = all.slice(1);
  const mandanteHorarioSourceMap = buildMandanteHorarioSourceMap_(dataRows, mIdx, hdr, obsI);

  const out    = [];
  const colors = [];

  dataRows.forEach(row => {
    const scheduleRow = getHorarioSourceRow_(row, mandanteHorarioSourceMap, mIdx, hdr, obsI);

    const area = row[mIdx['AREA']];
    const scheduleArea = scheduleRow[mIdx['AREA']];
    const areaNorm = String(scheduleArea || area || '').trim().toUpperCase();
    const plan = row[mIdx['PLAN DE ESTUDIO']];
    const nrc  = row[mIdx['NRC']];
    const scheduleNrc = scheduleRow[mIdx['NRC']];
    const liga = row[mIdx['CONECTOR DE LIGA']] || '';
    const lc   = row[mIdx['LC']];

    const split = splitCodigo(row[mIdx['CODIGO']], row[mIdx['MATERIA']], row[mIdx['CURSO']]);
    const mat   = split.materia;
    const cur   = split.curso;
    const cursoOut = cur ? "'" + cur : '';

    const sec  = row[mIdx['SECCIONES']];
    const tit  = normalizeName(row[mIdx['TITULO']]);

    const b1   = normalizeName(profFields.nombre1 >= 0 ? scheduleRow[profFields.nombre1] : '');
    const b2   = normalizeName(profFields.nombre2 >= 0 ? scheduleRow[profFields.nombre2] : '');
    const lab  = normalizeName(profFields.nombreLabt >= 0 ? scheduleRow[profFields.nombreLabt] : '');

    const hasClasesProgramar = isProgramarActive_(scheduleRow[mIdx['Clases A PROGRAMAR']]);
    const hasAyudProgramar   = isProgramarActive_(scheduleRow[mIdx['Ayudantías PROGRAMAR']]);
    const hasLabtProgramar   = isProgramarActive_(scheduleRow[mIdx['Laboratorios o Talleres PROGRAMAR']]);

    const onlyLabtProgramado = isOnlyLabtEffective_(
      scheduleRow,
      mIdx,
      days,
      hasClasesProgramar,
      hasAyudProgramar,
      hasLabtProgramar
    );
    const mainProf = b1 + (b2 ? '+' + b2 : '');

    const planComunValue = String(row[mIdx['Plan Común']] || '').trim();
    const isFirstSemester = planComunValue === '1';

    days.forEach((d, di) => {
      const cell = scheduleRow[mIdx[d]];
      if (!cell) return;

      cell.toString().split(',').forEach(part => {
        let [tipo, time] = part.trim().split(' ');
        tipo = tipo ? tipo.toUpperCase() : '';
        if (tipo === 'LABT') tipo = 'LAB/TALLER';

        const dayCells = Array(5).fill('');
        dayCells[di] = time;

        let ini = '', fin = '';

        if (tipo === 'CLAS') {
          ini = classStart;
          fin = classEnd;
        } else if (tipo === 'AYUD') {
          ini = ayudStart;
          fin = ayudEnd;
        } else if (tipo === 'LAB/TALLER') {
          ini = onlyLabtProgramado ? classStart : ayudStart;
          fin = ayudEnd;
        }

        if (ini && isFirstSemester) {
          ini = shiftDateString_(ini, -2, tz);
        }

        if (tipo === 'CLAS' && areaNorm === 'FACULTAD CEG') {
          const examDate = facultadCegExamMap[String(scheduleNrc || nrc || '').trim()];
          if (examDate) {
            const weekdayIndex = di + 1;
            const lastClassDate = previousWeekdayBeforeDate_(examDate, weekdayIndex);
            fin = formatDateDDMMYYYY_(lastClassDate, tz);
          }
        }

        let profOut;
        if (tipo === 'CLAS') {
          profOut = mainProf;
        } else if (tipo === 'LAB/TALLER') {
          profOut = lab ? lab : (onlyLabtProgramado ? mainProf : '');
        } else {
          profOut = b1;
        }

        out.push([
          area,plan,nrc,liga,lc,mat,cursoOut,sec,tit,
          ...dayCells, ini,fin,'', tipo, profOut
        ]);
      });
    });

    for (let c = obsI + 1; c < hdr.length; c++) {
      const dateObj = hdr[c];
      if (!(dateObj instanceof Date)) continue;

      const cell = scheduleRow[c];
      if (!cell) continue;

      const dateStr = Utilities.formatDate(dateObj, tz, 'dd/MM/yyyy');
      // dateStr is used as both INICIO and FIN for PRBA / EXAM rows.

      cell.toString().split(',').forEach(part => {
        let [tipo2, times] = part.trim().split(' ');
        tipo2 = tipo2 ? tipo2.toUpperCase() : '';
        if (tipo2 === 'LABT') tipo2 = 'LAB/TALLER';

        const [t0, t1] = times ? times.split('-') : ['', ''];

        let tag, bg;
        if (/CLAS|AYUD|LAB\/TALLER/.test(tipo2)) {
          tag = 'PRBA'; bg = '#CCEEFF';
        } else if (/TARDE/.test(tipo2)) {
          tag = 'PRBA'; bg = '#FCE4D6';
        } else if (/EXAM/.test(tipo2)) {
          tag = 'EXAM'; bg = '#FCE4D6';
        } else {
          tag = 'PRBA'; bg = '#CCEEFF';
        }

        const wd = dateObj.getDay();
        const dayCells = Array(5).fill('');
        if (wd >= 1 && wd <= 5) {
          dayCells[wd-1] = `${t0}-${t1}`;
        }

        let profOut2;
        if (tipo2 === 'CLAS') {
          profOut2 = mainProf;
        } else if (tipo2 === 'LAB/TALLER') {
          profOut2 = lab ? lab : (onlyLabtProgramado ? mainProf : '');
        } else {
          profOut2 = b1;
        }

        if (!profOut2 && wd >= 1 && wd <= 5) {
          const dayIndex = wd - 1;
          const dayName  = days[dayIndex];
          const weeklyCell = scheduleRow[mIdx[dayName]];

          if (weeklyCell) {
            const weeklyParts = weeklyCell.toString().split(',');
            const evalSlot = `${t0}-${t1}`;

            for (let pW of weeklyParts) {
              let [tipoW, timeW] = pW.trim().split(' ');
              let tipoWNorm = tipoW ? tipoW.toUpperCase() : '';
              if (tipoWNorm === 'LABT') tipoWNorm = 'LAB/TALLER';

              if (tipoWNorm === 'LAB/TALLER' && timeW === evalSlot) {
                profOut2 = onlyLabtProgramado ? mainProf : (b1 + (b2 ? '+' + b2 : ''));
                break;
              }
            }
          }
        }

        colors.push({ row: tableHeaderRow + out.length + 1, color: bg });
        // For PRBA / EXAM rows, both INICIO and FIN must be the evaluation date.
        out.push([
          area, plan, nrc, liga, lc, mat, cursoOut, sec, tit,
          ...dayCells, dateStr, dateStr, '', tag, profOut2
        ]);
      });
    }
  });

  if (out.length) {
    const startRow = tableHeaderRow + 1;
    horario.getRange(startRow, 1, out.length, HEADER.length)
           .setValues(out)
           .setBackgrounds(out.map(_ => Array(HEADER.length).fill(null)));
    colors.forEach(ci => {
      horario.getRange(ci.row, 1, 1, HEADER.length).setBackground(ci.color);
    });
    horario.getRange(startRow, 7, out.length, 1).setNumberFormat('@');
  }

  const periodSheet = ss.getSheetByName('NRC POR PERIODO');
  if (!periodSheet) {
    normalizeHorarioDateColumns_(horario);
    ui.alert('✅ HORARIO generado usando profesores desde MAESTRO. No se encontró la hoja NRC POR PERIODO, por lo que no se aplicó la sincronización complementaria de profesores.');
    return;
  }

  const pd2  = periodSheet.getDataRange().getValues();
  if (pd2.length < 2) {
    normalizeHorarioDateColumns_(horario);
    ui.alert('✅ HORARIO generado usando profesores desde MAESTRO. La hoja NRC POR PERIODO no tiene datos para sincronización complementaria de profesores.');
    return;
  }

  const pHdr = pd2[0].map(h => h.toString().toUpperCase());
  const pIdx = {};
  pHdr.forEach((h,i) => pIdx[h] = i);

  const profMap = {};
  const nrcSet  = new Set();

  pd2.slice(1).forEach(r => {
    const nrcCell = r[pIdx['NRC']];
    if (nrcCell === '' || nrcCell == null) return;
    const key = String(nrcCell).trim();
    if (!key) return;

    nrcSet.add(key);

    const prof   = r[pIdx['PROFESOR']];
    const profStr = prof == null ? '' : String(prof).trim();
    if (profStr) profMap[key] = profStr;
  });

  const horarioBlock = getHorarioBlock_(horario);
  const hd   = horarioBlock.values;
  const hIdx = horarioBlock.idx;
  const profCol = hIdx['PROFESOR'] + 1;

  for (let i = 1; i < hd.length; i++) {
    const row = hd[i];
    const sheetRow = horarioBlock.headerRow + i;
    const nrcKey = String(row[hIdx['NRC']] || '').trim();

    const origRaw = row[hIdx['PROFESOR']];
    const origStr = String(origRaw || '').trim();
    let [, suffix = ''] = origStr ? origStr.split('+') : ['',''];

    const tipoReu = String(row[hIdx['TIPO DE REUNION']] || '').toUpperCase();

    const hasNRC    = nrcSet.has(nrcKey);
    const mapped    = hasNRC ? profMap[nrcKey] : undefined;
    const hasMapped = typeof mapped === 'string' && mapped.trim() !== '';

    if (hasMapped) {
      if (tipoReu === 'LAB/TALLER') {
        continue;
      } else {
        const newVal = mapped + (suffix ? '+' + suffix : '');
        horario.getRange(sheetRow, profCol).setValue(newVal);
      }
    } else if (hasNRC) {
      horario.getRange(sheetRow, profCol).setBackground('#FFCCCC');
    } else {
      horario.getRange(sheetRow, 1, 1, HEADER.length).setBackground('#FFFF00');
    }
  }

  normalizeHorarioDateColumns_(horario);
  ui.alert('✅ HORARIO generado.');
}

function fillSalasHorario() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const horario = ss.getSheetByName('HORARIO');
  if (!horario) {
    ui.alert('No encontré la hoja HORARIO. Primero ejecuta “Crear Horario”.');
    return;
  }

  const pSheet = ss.getSheetByName('NRC POR PERIODO');
  if (!pSheet) {
    ui.alert('No encontré la hoja NRC POR PERIODO.');
    return;
  }

  const src = SpreadsheetApp.openById(getSourceSpreadsheetId_());
  const mSheet = src.getSheetByName('MAESTRO');
  if (!mSheet) {
    ui.alert('No encontré la hoja MAESTRO en el documento fuente.');
    return;
  }

  const tz = ss.getSpreadsheetTimeZone();

  function normalizeTimeRange_(s) {
    s = String(s || '').trim();
    if (!s) return '';
    const m = s.match(/^(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})$/);
    if (!m) return s.replace(/\s+/g, '');
    const h1 = String(Number(m[1])).padStart(2, '0');
    const h2 = String(Number(m[3])).padStart(2, '0');
    return `${h1}:${m[2]}-${h2}:${m[4]}`;
  }

  let horarioBlock;
  try {
    horarioBlock = getHorarioBlock_(horario);
  } catch (e) {
    ui.alert(e.message);
    return;
  }

  const hVals = horarioBlock.values;
  const hIdx  = horarioBlock.idx;
  const dataStartRow = horarioBlock.headerRow + 1;
  const dataRowCount = Math.max(0, hVals.length - 1);

  const requiredHorario = ['NRC', 'TIPO DE REUNION', 'INICIO', 'FIN', 'SALA', 'CURSO'];
  for (const c of requiredHorario) {
    if (!(c in hIdx)) {
      ui.alert(`Falta la columna "${c}" en HORARIO.`);
      return;
    }
  }

  const existingBackgrounds = dataRowCount > 0
    ? horario.getRange(dataStartRow, 1, dataRowCount, horarioBlock.lastCol).getBackgrounds()
    : [];

  const pVals = pSheet.getDataRange().getValues();
  if (pVals.length < 2) {
    ui.alert('NRC POR PERIODO no tiene datos.');
    return;
  }

  const pHdrU = pVals[0].map(h => String(h).trim().toUpperCase());
  const pIdx = {};
  pHdrU.forEach((h, i) => pIdx[h] = i);

  const requiredPeriodo = ['NRC', 'TIPO', 'SALA', 'INICIO', 'FIN'];
  for (const c of requiredPeriodo) {
    if (!(c in pIdx)) {
      ui.alert(`Falta la columna "${c}" en NRC POR PERIODO.`);
      return;
    }
  }

  const cegExamMap = buildFacultadCegExamMap_(mSheet, tz);

  const periodMap = {};

  for (let r = 1; r < pVals.length; r++) {
    const row = pVals[r];

    const nrc = String(row[pIdx['NRC']] || '').trim();
    const tipo = getTipoNormalized_(row[pIdx['TIPO']]);
    const sala = String(row[pIdx['SALA']] || '').trim();

    if (!nrc || !tipo) continue;

    const dayInfo = getDayColumnNrcPeriodo_(row, pIdx);
    const day = String(dayInfo.day || '').trim().toUpperCase();
    const timeRaw = String(dayInfo.time || '').trim();
    const timeNorm = normalizeTimeRange_(timeRaw);

    const ini = parseDateFlexible_(row[pIdx['INICIO']], tz);
    const fin = parseDateFlexible_(row[pIdx['FIN']], tz);
    if (!ini || !fin) continue;

    const key = [nrc, tipo, day].join('|');
    if (!periodMap[key]) periodMap[key] = [];

    periodMap[key].push({
      timeRaw,
      timeNorm,
      start: ini,
      end: fin,
      sala
    });
  }

  Object.keys(periodMap).forEach(key => {
    const grouped = {};
    periodMap[key].forEach(it => {
      const gk = `${it.sala || ''}|||${it.timeNorm || ''}`;
      if (!grouped[gk]) grouped[gk] = [];
      grouped[gk].push(it);
    });

    const mergedAll = [];
    Object.keys(grouped).forEach(gk => {
      const arr = grouped[gk];
      const merged = mergeIntervals_(
        arr.map(x => ({
          start: x.start,
          end: x.end,
          sala: x.sala
        }))
      );
      const timeRaw = arr[0].timeRaw || '';
      const timeNorm = arr[0].timeNorm || '';
      merged.forEach(m => {
        mergedAll.push({
          timeRaw,
          timeNorm,
          start: m.start,
          end: m.end,
          sala: m.sala
        });
      });
    });

    periodMap[key] = mergedAll.sort((a, b) => a.start - b.start || a.end - b.end);
  });

  const HEADER = horarioBlock.headers.slice();
  const outRows = [];
  const outBackgrounds = [];

  let splitCount = 0;
  let filledCount = 0;
  let ambiguousCount = 0;
  let timeMismatchCount = 0;

  for (let i = 1; i < hVals.length; i++) {
    const row = hVals[i];
    const originalBg = existingBackgrounds[i - 1] || Array(HEADER.length).fill(null);

    const tipo = getTipoNormalized_(row[hIdx['TIPO DE REUNION']]);
    const nrc = String(row[hIdx['NRC']] || '').trim();

    const dayName = getDayNameFromHorarioRow_(row, hIdx);
    const dayUpper = dayName ? dayName.toUpperCase() : '';
    const timeNorm = normalizeTimeRange_(getTimeFromHorarioRow_(row, hIdx));

    const originalIni = parseDateFlexible_(row[hIdx['INICIO']], tz);
    const originalFin = parseDateFlexible_(row[hIdx['FIN']], tz);

    const isRecurring = ['CLAS', 'AYUD', 'LAB/TALLER'].includes(tipo);

    if (!isRecurring || !nrc || !dayUpper || !originalIni || !originalFin) {
      outRows.push(row.slice());
      outBackgrounds.push(originalBg.slice());
      continue;
    }

    let effectiveFin = new Date(originalFin.getTime());

    if (tipo === 'CLAS' && cegExamMap[nrc] && dayUpper) {
      const dayIndexMap = {
        LUNES: 1,
        MARTES: 2,
        MIERCOLES: 3,
        JUEVES: 4,
        VIERNES: 5
      };

      const weekdayIndex = dayIndexMap[dayUpper];
      if (weekdayIndex) {
        const lastClassDate = previousWeekdayBeforeDate_(cegExamMap[nrc], weekdayIndex);
        effectiveFin = minDate_(effectiveFin, lastClassDate);
      }
    }

    if (effectiveFin.getTime() < originalIni.getTime()) {
      const newRow = row.slice();
      newRow[hIdx['INICIO']] = formatDateDDMMYYYY_(originalIni, tz);
      newRow[hIdx['FIN']] = formatDateDDMMYYYY_(effectiveFin, tz);
      newRow[hIdx['SALA']] = '';
      outRows.push(newRow);
      outBackgrounds.push(originalBg.slice());
      continue;
    }

    const key = [nrc, tipo, dayUpper].join('|');
    const candidates = periodMap[key] || [];

    // La SALA solo puede heredarse cuando coinciden NRC, tipo, día Y horario.
    // Antes, si existía un único candidato para NRC+tipo+día, se aceptaba aunque
    // el intervalo fuera distinto. Eso podía asignar, por ejemplo, una sala de
    // AYUD 15:30-17:20 a una fila AYUD 13:30-15:20.
    let matched = candidates.filter(c => c.timeNorm === timeNorm);

    if (!matched.length) {
      const newRow = row.slice();
      newRow[hIdx['INICIO']] = formatDateDDMMYYYY_(originalIni, tz);
      newRow[hIdx['FIN']] = formatDateDDMMYYYY_(effectiveFin, tz);

      // Limpia cualquier asignación previa o incorrecta. Si no existe coincidencia
      // exacta de día y hora en NRC POR PERIODO, la SALA debe quedar vacía.
      newRow[hIdx['SALA']] = '';

      outRows.push(newRow);
      outBackgrounds.push(originalBg.slice());

      if (candidates.length) {
        timeMismatchCount++;
        if (candidates.length > 1) ambiguousCount++;
      }
      continue;
    }

    matched = matched
      .map(c => ({
        start: maxDate_(c.start, originalIni),
        end: minDate_(c.end, effectiveFin),
        sala: c.sala
      }))
      .filter(c => c.start.getTime() <= c.end.getTime());

    if (!matched.length) {
      const newRow = row.slice();
      newRow[hIdx['INICIO']] = formatDateDDMMYYYY_(originalIni, tz);
      newRow[hIdx['FIN']] = formatDateDDMMYYYY_(effectiveFin, tz);
      outRows.push(newRow);
      outBackgrounds.push(originalBg.slice());
      continue;
    }

    matched = mergeIntervals_(matched);

    if (matched.length === 1) {
      const newRow = row.slice();
      newRow[hIdx['INICIO']] = formatDateDDMMYYYY_(matched[0].start, tz);
      newRow[hIdx['FIN']] = formatDateDDMMYYYY_(matched[0].end, tz);
      newRow[hIdx['SALA']] = matched[0].sala;
      outRows.push(newRow);
      outBackgrounds.push(originalBg.slice());
      if (matched[0].sala) filledCount++;
    } else {
      splitCount++;
      matched.forEach(intv => {
        const newRow = row.slice();
        newRow[hIdx['INICIO']] = formatDateDDMMYYYY_(intv.start, tz);
        newRow[hIdx['FIN']] = formatDateDDMMYYYY_(intv.end, tz);
        newRow[hIdx['SALA']] = intv.sala;
        outRows.push(newRow);
        outBackgrounds.push(originalBg.slice());
        if (intv.sala) filledCount++;
      });
    }
  }

  const lastRow = horario.getLastRow();
  if (lastRow >= dataStartRow) {
    horario.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, horarioBlock.lastCol)
      .clearContent()
      .clearDataValidations();
  }

  if (outRows.length) {
    horario.getRange(dataStartRow, 1, outRows.length, HEADER.length).setValues(outRows);
    horario.getRange(dataStartRow, 1, outRows.length, HEADER.length).setBackgrounds(outBackgrounds);
    horario.getRange(dataStartRow, hIdx['CURSO'] + 1, outRows.length, 1).setNumberFormat('@');
  }

  ui.alert(
    `✅ SALAS asignadas / breaks procesados.\n` +
    `Filas con split: ${splitCount}\n` +
    `Filas con sala asignada: ${filledCount}\n` +
    `Filas ambiguas sin cambio: ${ambiguousCount}\n` +
    `Filas sin sala por diferencia de horario: ${timeMismatchCount}`
  );
}

function updateSalasPruebas() {
  const ui      = SpreadsheetApp.getUi();
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const maestro = SpreadsheetApp.openById(getSourceSpreadsheetId_()).getSheetByName('MAESTRO');
  const horario = ss.getSheetByName('HORARIO');
  const pSheet  = ss.getSheetByName('NRC POR PERIODO');

  if (!maestro || !horario || !pSheet) {
    ui.alert('Asegúrate de que existan MAESTRO, HORARIO y NRC POR PERIODO.');
    return;
  }

  const tz = ss.getSpreadsheetTimeZone();

  function normalizeTimeRangeLocal_(s) {
    s = String(s || '').trim();
    if (!s) return '';

    const m = s.match(/^(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})$/);
    if (!m) return s.replace(/\s+/g, '');

    const h1 = String(Number(m[1])).padStart(2, '0');
    const h2 = String(Number(m[3])).padStart(2, '0');
    return `${h1}:${m[2]}-${h2}:${m[4]}`;
  }

  function pushMap_(map, key, item) {
    if (!key) return;
    if (!map[key]) map[key] = [];
    map[key].push(item);
  }

  function dateInsideInterval_(dateObj, startObj, endObj) {
    if (!dateObj || !startObj || !endObj) return false;
    return dateObj.getTime() >= startObj.getTime() &&
           dateObj.getTime() <= endObj.getTime();
  }

  function getExactSlotCandidates_(candidates, targetDate, targetDayUpper, targetTimeNorm) {
    return (candidates || []).filter(c => {
      if (!dateInsideInterval_(targetDate, c.start, c.end)) return false;

      if (targetDayUpper && c.dayUpper && c.dayUpper !== targetDayUpper) return false;
      if (targetTimeNorm && c.timeNorm && c.timeNorm !== targetTimeNorm) return false;

      return true;
    });
  }

  function firstSalaFromCandidates_(candidates) {
    for (const c of candidates || []) {
      const sala = String(c.sala || '').trim();
      if (sala) return sala;
    }
    return '';
  }

  function uniqueSalaFromTypedInterval_(candidates, targetDate) {
    const salas = [];
    for (const c of candidates || []) {
      if (!dateInsideInterval_(targetDate, c.start, c.end)) continue;
      const sala = String(c.sala || '').trim();
      if (sala && !salas.includes(sala)) salas.push(sala);
    }
    return salas.length === 1 ? salas[0] : '';
  }

  // --------------------------------------------------------------------
  // 1) MAESTRO: NRC|fecha -> tipo real (CLAS / AYUD / LAB/TALLER)
  //    Esto sirve como apoyo, pero ya no es la única forma de resolver PRBA.
  //    En filas con LC, las PRBA pueden haber sido heredadas desde el curso
  //    mandante, por lo que el tipo real puede no existir para el NRC mostrado.
  // --------------------------------------------------------------------
  const mData = maestro.getDataRange().getValues();
  if (mData.length < 2) {
    ui.alert('MAESTRO no tiene datos.');
    return;
  }

  const mHdr  = mData[0];
  const mIdx  = {};
  mHdr.forEach((h, i) => mIdx[h] = i);

  if (!('NRC' in mIdx) || !('OBSERVACION' in mIdx)) {
    ui.alert('Faltan columnas NRC u OBSERVACION en MAESTRO.');
    return;
  }

  const obsI  = mIdx['OBSERVACION'];
  const tipoRealPorFecha = {};

  for (let r = 1; r < mData.length; r++) {
    const row = mData[r];
    const nrc = String(row[mIdx['NRC']] || '').trim();
    if (!nrc) continue;

    for (let c = obsI + 1; c < mHdr.length; c++) {
      const headerDate = mHdr[c];
      if (!(headerDate instanceof Date)) continue;

      const cell = row[c];
      if (!cell) continue;

      const fechaStr = Utilities.formatDate(headerDate, tz, 'dd/MM/yyyy');
      const parts = String(cell).split(',');

      for (const rawPart of parts) {
        const part = String(rawPart || '').trim().toUpperCase();
        if (!part) continue;

        let tipo = part.split(' ')[0];
        if (tipo === 'LABT') tipo = 'LAB/TALLER';

        if (['CLAS', 'AYUD', 'LAB/TALLER'].includes(tipo)) {
          tipoRealPorFecha[`${nrc}|${fechaStr}`] = tipo;
        }
      }
    }
  }

  // --------------------------------------------------------------------
  // 2) HORARIO
  // --------------------------------------------------------------------
  let horarioBlock;
  try {
    horarioBlock = getHorarioBlock_(horario);
  } catch (e) {
    ui.alert(e.message);
    return;
  }

  const hData = horarioBlock.values;
  const hIdx  = horarioBlock.idx;

  const requiredHorario = ['NRC', 'TIPO DE REUNION', 'INICIO', 'FIN', 'SALA'];
  for (const col of requiredHorario) {
    if (!(col in hIdx)) {
      ui.alert(`Falta la columna "${col}" en HORARIO.`);
      return;
    }
  }

  const hasLcColumn = ('LC' in hIdx);

  // --------------------------------------------------------------------
  // 3) NRC POR PERIODO
  //    Índice directo para PRBA / EXAM.
  //    key = NRC|TIPO|DAY -> [{timeNorm,start,end,sala}]
  // --------------------------------------------------------------------
  const pData = pSheet.getDataRange().getValues();
  if (pData.length < 2) {
    ui.alert('NRC POR PERIODO no tiene datos.');
    return;
  }

  const pHdr = pData[0].map(h => String(h).trim().toUpperCase());
  const pIdx = {};
  pHdr.forEach((h, i) => pIdx[h] = i);

  const requiredPeriodo = ['NRC', 'TIPO', 'SALA', 'INICIO', 'FIN'];
  for (const col of requiredPeriodo) {
    if (!(col in pIdx)) {
      ui.alert(`Falta la columna "${col}" en NRC POR PERIODO.`);
      return;
    }
  }

  const directEvalMap = {}; // NRC|TIPO|DAY -> [{timeNorm,start,end,sala}]
  for (let r = 1; r < pData.length; r++) {
    const row = pData[r];
    const nrc = String(row[pIdx['NRC']] || '').trim();
    let tipo  = String(row[pIdx['TIPO']] || '').trim().toUpperCase();
    if (!nrc || !tipo) continue;

    if (tipo !== 'PRBA' && tipo !== 'EXAM') continue;

    const sala = String(row[pIdx['SALA']] || '').trim();
    const ini  = parseDateFlexible_(row[pIdx['INICIO']], tz);
    const fin  = parseDateFlexible_(row[pIdx['FIN']], tz);
    if (!ini || !fin) continue;

    const dayInfo = getDayColumnNrcPeriodo_(row, pIdx);
    const day  = String(dayInfo.day || '').trim().toUpperCase(); // puede venir vacío
    const timeNorm = normalizeTimeRangeLocal_(dayInfo.time);

    const key = `${nrc}|${tipo}|${day}`;
    pushMap_(directEvalMap, key, {
      timeNorm,
      start: ini,
      end: fin,
      sala
    });
  }

  // --------------------------------------------------------------------
  // 4) Índices de filas recurrentes de HORARIO para fallback de PRBA.
  //
  //    Antes se dependía casi exclusivamente de:
  //      tipoRealPorFecha[NRC|fecha]
  //    Eso falla en cursos con LC, porque la PRBA puede estar heredada desde
  //    el curso mandante y no existir como observación propia del NRC mostrado.
  //
  //    Nueva lógica:
  //    - Buscar la fila recurrente del mismo NRC que coincida por día + hora.
  //    - Solo si no existe esa fila, probar por LC.
  //    - Si la fila recurrente existe pero no tiene SALA, se deja sin SALA.
  // --------------------------------------------------------------------
  const recurringByNrc = {};
  const recurringByNrcTipo = {};
  const recurringByLc = {};

  for (let i = 1; i < hData.length; i++) {
    const row = hData[i];
    const nrc = String(row[hIdx['NRC']] || '').trim();
    const lc  = hasLcColumn ? String(row[hIdx['LC']] || '').trim() : '';

    let tipo = String(row[hIdx['TIPO DE REUNION']] || '').trim().toUpperCase();
    if (tipo === 'LABT') tipo = 'LAB/TALLER';

    if (!['CLAS', 'AYUD', 'LAB/TALLER'].includes(tipo)) continue;
    if (!nrc) continue;

    const ini = parseDateFlexible_(row[hIdx['INICIO']], tz);
    const fin = parseDateFlexible_(row[hIdx['FIN']], tz);
    if (!ini || !fin) continue;

    const dayName = getDayNameFromHorarioRow_(row, hIdx);
    const dayUpper = dayName ? dayName.toUpperCase() : '';
    const timeNorm = normalizeTimeRangeLocal_(getTimeFromHorarioRow_(row, hIdx));
    const sala = String(row[hIdx['SALA']] || '').trim();

    const item = {
      nrc,
      lc,
      tipo,
      dayUpper,
      timeNorm,
      start: ini,
      end: fin,
      sala
    };

    pushMap_(recurringByNrc, nrc, item);
    pushMap_(recurringByNrcTipo, `${nrc}|${tipo}`, item);
    if (lc) pushMap_(recurringByLc, lc, item);
  }

  Object.keys(recurringByNrc).forEach(key => {
    recurringByNrc[key].sort((a, b) => a.start - b.start || a.end - b.end);
  });
  Object.keys(recurringByNrcTipo).forEach(key => {
    recurringByNrcTipo[key].sort((a, b) => a.start - b.start || a.end - b.end);
  });
  Object.keys(recurringByLc).forEach(key => {
    recurringByLc[key].sort((a, b) => a.start - b.start || a.end - b.end);
  });

  // --------------------------------------------------------------------
  // 5) Resolver PRBA / EXAM
  // --------------------------------------------------------------------
  let updatedDirect = 0;
  let updatedFallbackSameNrc = 0;
  let updatedFallbackLc = 0;
  let missingDirect = 0;
  let missingFallback = 0;
  let legitimateNoSala = 0;

  for (let i = 1; i < hData.length; i++) {
    const row = hData[i];
    const tipoHorario = String(row[hIdx['TIPO DE REUNION']] || '').trim().toUpperCase();
    if (tipoHorario !== 'PRBA' && tipoHorario !== 'EXAM') continue;

    const nrc = String(row[hIdx['NRC']] || '').trim();
    const lc  = hasLcColumn ? String(row[hIdx['LC']] || '').trim() : '';
    if (!nrc) continue;

    const ini = parseDateFlexible_(row[hIdx['INICIO']], tz);
    if (!ini) continue;

    const fechaStr = Utilities.formatDate(ini, tz, 'dd/MM/yyyy');
    const dayName = getDayNameFromHorarioRow_(row, hIdx);
    const dayUpper = dayName ? dayName.toUpperCase() : '';
    const timeNorm = normalizeTimeRangeLocal_(getTimeFromHorarioRow_(row, hIdx));

    const sheetRow = horarioBlock.headerRow + i;
    let salaAsignada = '';

    // --------------------------------------------------------------
    // A) Intento directo desde NRC POR PERIODO para PRBA / EXAM
    // --------------------------------------------------------------
    const directKey = `${nrc}|${tipoHorario}|${dayUpper}`;
    let directCandidates = directEvalMap[directKey] || [];

    // Si no hay con día, probar sin día.
    if (!directCandidates.length) {
      const blankDayKey = `${nrc}|${tipoHorario}|`;
      directCandidates = directEvalMap[blankDayKey] || [];
    }

    // Preferir coincidencia exacta de horario.
    let directMatched = directCandidates.filter(c =>
      !timeNorm || !c.timeNorm || c.timeNorm === timeNorm
    );

    if (!directMatched.length) {
      directMatched = directCandidates.slice();
    }

    for (const cand of directMatched) {
      if (dateInsideInterval_(ini, cand.start, cand.end)) {
        salaAsignada = String(cand.sala || '').trim();
        if (salaAsignada) break;
      }
    }

    if (salaAsignada) {
      horario.getRange(sheetRow, hIdx['SALA'] + 1).setValue(salaAsignada);
      updatedDirect++;
      continue;
    }

    // EXAM sin match directo: no se infiere desde CLAS/AYUD/LAB.
    if (tipoHorario === 'EXAM') {
      missingDirect++;
      continue;
    }

    // --------------------------------------------------------------
    // B) Fallback PRBA por día + hora usando las filas recurrentes
    //    ya existentes en HORARIO.
    //
    //    Esto soluciona los casos con LC:
    //    Si la PRBA está en Miércoles 15:30-17:20, busca la AYUD/CLAS/LAB
    //    del mismo NRC en Miércoles 15:30-17:20 y toma su SALA.
    // --------------------------------------------------------------
    const sameNrcCandidates = recurringByNrc[nrc] || [];
    const sameNrcExact = getExactSlotCandidates_(
      sameNrcCandidates,
      ini,
      dayUpper,
      timeNorm
    );

    if (sameNrcExact.length) {
      salaAsignada = firstSalaFromCandidates_(sameNrcExact);

      if (salaAsignada) {
        horario.getRange(sheetRow, hIdx['SALA'] + 1).setValue(salaAsignada);
        updatedFallbackSameNrc++;
        continue;
      }

      // Importante:
      // Si la fila recurrente del mismo NRC existe pero no tiene SALA,
      // se considera un caso legítimo sin sala y NO se hereda desde otro LC.
      legitimateNoSala++;
      missingFallback++;
      missingDirect++;
      continue;
    }

    // --------------------------------------------------------------
    // C) Fallback adicional usando tipo real desde MAESTRO,
    //    cuando existe y no hubo coincidencia exacta por día + hora.
    // --------------------------------------------------------------
    const tipoReal = tipoRealPorFecha[`${nrc}|${fechaStr}`];
    if (tipoReal) {
      const typedCandidates = recurringByNrcTipo[`${nrc}|${tipoReal}`] || [];
      salaAsignada = uniqueSalaFromTypedInterval_(typedCandidates, ini);

      if (salaAsignada) {
        horario.getRange(sheetRow, hIdx['SALA'] + 1).setValue(salaAsignada);
        updatedFallbackSameNrc++;
        continue;
      }
    }

    // --------------------------------------------------------------
    // D) Fallback por LC solo si NO existe una fila recurrente exacta
    //    para ese NRC. Esto evita asignar salas por error cuando el
    //    propio horario del NRC tiene el bloque, pero sin sala.
    // --------------------------------------------------------------
    if (lc) {
      const lcCandidates = recurringByLc[lc] || [];
      const lcExact = getExactSlotCandidates_(
        lcCandidates,
        ini,
        dayUpper,
        timeNorm
      );

      salaAsignada = firstSalaFromCandidates_(lcExact);

      if (salaAsignada) {
        horario.getRange(sheetRow, hIdx['SALA'] + 1).setValue(salaAsignada);
        updatedFallbackLc++;
        continue;
      }
    }

    missingFallback++;
    missingDirect++;
  }

  // Mantener INICIO/FIN como texto dd/MM/yyyy aunque alguna operación previa
  // haya dejado objetos Date en esas columnas.
  if (typeof normalizeHorarioDateColumns_ === 'function') {
    normalizeHorarioDateColumns_(horario);
  }

  ui.alert(
    `✅ Salas actualizadas\n` +
    `Directas (PRBA/EXAM): ${updatedDirect}\n` +
    `Fallback PRBA mismo NRC por día/hora: ${updatedFallbackSameNrc}\n` +
    `Fallback PRBA por LC: ${updatedFallbackLc}\n` +
    `Sin sala legítima por bloque recurrente vacío: ${legitimateNoSala}\n` +
    `Sin sala encontrada: ${missingDirect}`
  );
}
