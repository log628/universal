/***************************************************************
 * getFORM — DUAL (OZ/WB)
 * Универсальный запуск layout без кулдаунов:
 *   runLayoutImmediate(selectedCab?)
 * Источники данных выбираются строго по платформе из ⚙️ Параметры!I2:
 *   - [OZ]/[WB] Артикулы  (A:M, 13 колонок; M = «Своя категория»)
 *   - [OZ]/[WB] Физ. оборот (A:G)
 ***************************************************************/

//////////////////// Константы ////////////////////
// — всё берём строго из REF, без фолбеков
const SHEET_CALC   = REF.SHEETS.CALC;
const SHEET_PARAMS = REF.SHEETS.PARAMS;

// — источники листов
const ARTS_OZ  = REF.SHEETS.ARTS_OZ;
const ARTS_WB  = REF.SHEETS.ARTS_WB;
const PHYS_OZ  = REF.SHEETS.FIZ_OZ;
const PHYS_WB  = REF.SHEETS.FIZ_WB;

// — единый контрол выбора кабинета
const CTRL_RANGE_A1 = REF.CTRL_RANGE_A1;

const ROW_DATA       = 4;
const MIN_DATA_ROWS  = 22;
const MIN_LAST_ROW   = ROW_DATA + MIN_DATA_ROWS - 1;

const DIVIDERS = [9, 12, 17, 20, 26]; // I, L, Q, T, Z
const WIDTHS = {
  R: 57, Q: 104, T: 104,
  narrow62: ['J','K','M','N','O','P'],
  other85:  ['H','V','W','X','Y','AA','AB'],
  separators: 3
};

const FONT  = { data: 10, family: 'Roboto' };
const COLOR = { txt: '#000000', inner: '#b7b7b7', outer: '#000000', white:'#ffffff' };
const PALETTE = {
  introF: '#efefef',
  flow:   '#fff2cc',
  calcT:  '#e9e2f8',
  profit: '#fce5cd'
};

const PARAMS_MODE_KEY = 'ключи'; // пока режим G не используется, но оставим для совместимости


/********************* ПУБЛИЧНЫЕ ХЕНДЛЕРЫ ************************/

/** МГНОВЕННЫЙ запуск рендера текущей площадки (без кулдаунов/локов) */
function runLayoutImmediate(selectedCab) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_CALC);
  if (!sh) throw new Error(`Лист "${SHEET_CALC}" не найден`);

  const ctrl = sh.getRange(CTRL_RANGE_A1);
  const currentCab = String(selectedCab || ctrl.getDisplayValue() || '').trim();
  if (!currentCab) {
    ss.toast('Не выбран кабинет (контрол ' + CTRL_RANGE_A1 + ')', 'Внимание', 3);
    return;
  }

  // На время рендера отключаем валидацию
  removeCabinetDropdown_(ctrl);
  try {
    // Рендер калькулятора (площадка по ⚙️Параметры!I2)
    if (typeof layoutCalculator === 'function') {
      layoutCalculator(currentCab);
    }
    // «⛓️ Параллель» — тем же кабинетом/списком артикулов
    if (typeof layoutParallel === 'function') {
      layoutParallel(currentCab);
    }
    SpreadsheetApp.flush();
  } finally {
    // Возвращаем дропдаун и выбранное значение
    restoreCabinetDropdown_(ctrl, currentCab);
  }
}

/** onEdit: если меняется контрол кабинета — сразу рендерим */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (!sh || sh.getName() !== SHEET_CALC) return;

    const ctrl = sh.getRange(CTRL_RANGE_A1);
    if (!rangeIntersects_(e.range, ctrl)) return;

    const selectedCab = String(ctrl.getDisplayValue() || '').trim();
    if (!selectedCab) return;

    runLayoutImmediate(selectedCab);
  } catch (err) {
    throw err;
  }
}


/********************* КОНТРОЛ КАБИНЕТА **************************/

function setupCabinetControl_() {
  const ss = SpreadsheetApp.getActive();
  const shCalc = ss.getSheetByName(SHEET_CALC);
  if (!shCalc) throw new Error(`Лист "${SHEET_CALC}" не найден`);

  const ctrl = shCalc.getRange(CTRL_RANGE_A1);
  const currentValue = String(ctrl.getDisplayValue() || '').trim();

  ctrl.breakApart();
  ctrl.clearDataValidations();
  ctrl.merge();
  ctrl.setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontFamily(FONT.family)
      .setFontSize(11);

  restoreCabinetDropdown_(ctrl, currentValue || null);
}

function removeCabinetDropdown_(ctrlRange) {
  ctrlRange.clearDataValidations();
}

function restoreCabinetDropdown_(ctrlRange, selectedCab) {
  const list = getCabinetListFromParams_(); // уже фильтрует по ⚙️ Параметры!I2
  ctrlRange.clearDataValidations();

  if (!list.length) {
    if (selectedCab) ctrlRange.setValue(selectedCab);
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  ctrlRange.setDataValidation(rule);

  const cur = String(ctrlRange.getDisplayValue() || '').trim();
  const chosen =
    (selectedCab && list.indexOf(selectedCab) !== -1) ? selectedCab :
    (list.indexOf(cur) !== -1) ? cur :
    list[0];

  ctrlRange.setValue(chosen);
}

/** Возвращает список кабинетов из «⚙️ Параметры», с учётом фильтра I2 (OZON/WB/пусто) */
function getCabinetListFromParams_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_PARAMS);
  if (!sh) return [];

  const filterRaw = String(sh.getRange('I2').getDisplayValue() || '').trim();
  const filterUP  = filterRaw.toUpperCase();

  const last = sh.getLastRow();
  if (last < 2) return [];

  const rows = sh.getRange(2, 1, last - 1, 4).getDisplayValues(); // A..D

  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const name = String(rows[i][0] || '').trim();               // A Кабинет
    const plat = String(rows[i][3] || '').trim().toUpperCase(); // D Площадка
    if (!name) break; // до первой пустой строки

    if (filterUP) {
      if (plat === filterUP) out.push(name);
    } else {
      out.push(name);
    }
  }
  return Array.from(new Set(out));
}


/********************* ОСНОВНОЙ LAYOUT **************************/

function layoutCalculator(cabinetOverride) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_CALC);
  if (!sh) throw new Error(`Лист "${SHEET_CALC}" не найден`);

  // Единственный режим — «Артикулы»
  const src = collectRowsForCalculator_(cabinetOverride /* mode not used */);
  const rowsLen = Math.max(src.displayG.length, MIN_DATA_ROWS);

  const needLast = Math.max(ROW_DATA + rowsLen - 1, MIN_LAST_ROW);
  ensureColCapacityTo_(sh, 29);
  ensureRowCapacityTo_(sh, needLast);
  trimRowsAfter_(sh, rowsLen);

  clearPrimerAfterSizing_(sh, rowsLen);

  // G + заголовок + автоширина
  writeColumnG_(sh, src.displayG, rowsLen);
  setGHeader_(sh); // фикс «Артикул»
  autoWidthPlus_(sh, col_('G'), 50);

  // «Отзывы» (J:K) и «СС» (AA)
  writeReviews_(sh, src.ratingD, src.countC, rowsLen);
  writeSS_(sh, src.ssAA, rowsLen);

  // Блок M:P — M = только «Наличие»
  writeFlowBlock_(sh, src.flowM, src.flowN, src.flowO, src.flowP, rowsLen);

  applyNumberFormatsRUAB_(sh, rowsLen);

  applyWidths_(sh);
  applyDataFormatting_Only_(sh, rowsLen);
  applyDataBackgrounds_(sh, rowsLen);
  applyDataGrid_(sh);

  sh.getRange(ROW_DATA, col_('U'), rowsLen, 1).setFontWeight('bold');
}

/** ПЛАТФОРМА строго по ⚙️ Параметры!I2 → 'OZ' | 'WB' (дефолт 'OZ') */
function resolvePlatformCurrent_() {
  const ss = SpreadsheetApp.getActive();
  const shParams = ss.getSheetByName(SHEET_PARAMS);
  const raw = shParams ? String(shParams.getRange('I2').getDisplayValue() || '').trim().toUpperCase() : '';

  if (raw === 'OZON' || raw === 'OZ') return 'OZ';
  if (raw === 'WILDBERRIES' || raw === 'WB') return 'WB';

  try { SpreadsheetApp.getActive().toast('Площадка не указана в I2 — выбран OZ по умолчанию', 'Внимание', 3); } catch (_) {}
  return 'OZ';
}

function collectRowsForCalculator_(cabinetOverride /* mode not used */) {
  const ss  = SpreadsheetApp.getActive();

  // Текущий кабинет
  let selectedCab = String(cabinetOverride || '').trim();
  if (!selectedCab) {
    const shCalc = ss.getSheetByName(SHEET_CALC);
    if (shCalc) selectedCab = String(shCalc.getRange(CTRL_RANGE_A1).getDisplayValue() || '').trim();
  }
  if (!selectedCab) return emptyCalcRows_();

  // Площадка по I2
  const plat = resolvePlatformCurrent_(); // 'OZ' | 'WB'

  // Источники
  const artsSheetName = (plat === 'WB') ? ARTS_WB : ARTS_OZ;
  const physSheetName = (plat === 'WB') ? PHYS_WB : PHYS_OZ;

  // Чтение артикулов (A:M = 13 колонок, где M = «Своя категория»)
  const shS = ss.getSheetByName(artsSheetName);
  if (!shS) return emptyCalcRows_();

  const lastRow = shS.getLastRow();
  const lastCol = shS.getLastColumn();
  if (lastRow < 2 || lastCol < 13) return emptyCalcRows_();

  const headers = shS.getRange(1,1,1,13).getDisplayValues()[0];
  const colCab    = findHeaderIndexFlexible_(headers, ['Кабинет'])         || 1;  // A
  const colArt    = findHeaderIndexFlexible_(headers, ['Артикул'])         || 2;  // B
  const colRevsC  = findHeaderIndexFlexible_(headers, ['Отзывы'])          || 3;  // C
  const colRateD  = findHeaderIndexFlexible_(headers, ['Рейтинг'])         || 4;  // D
  const colPrice  = findHeaderIndexFlexible_(headers, ['Цена'])            || 10; // J (резерв на будущее)
  const colOwnCat = findHeaderIndexFlexible_(headers, ['Своя категория'])  || 13; // M

  const vals = shS.getRange(2,1,lastRow-1,13).getDisplayValues();

  // Фильтр по кабинету
  const filtered = vals.filter(row => {
    const cab  = String(row[colCab -1] || '').trim();
    const art  = String(row[colArt -1] || '').trim();
    return art && cab === selectedCab;
  });

  // Сортировка по Артикулу
  filtered.sort((a,b) => {
    const A = String(a[colArt-1]||'').trim();
    const B = String(b[colArt-1]||'').trim();
    return A.localeCompare(B, 'ru');
  });

  // «🍔 СС»!A:J → Map<tovar -> {cc,nal,vput,vpost}>
  const ssAJ = (REF && REF.readSS_AJ_Map) ? REF.readSS_AJ_Map() : new Map();

  // «Физ. оборот» → для N,O,P
  const physMap = readPhysMapForCabinet_(physSheetName);

  // Выход
  const displayG = [];
  const ratingD  = [];
  const countC   = [];
  const ssAA     = [];

  const flowM = []; // ← только «Наличие»
  const flowN = [];
  const flowO = [];
  const flowP = [];

  filtered.forEach(row => {
    const cab   = String(row[colCab   -1] || '').trim();
    const art   = String(row[colArt   -1] || '').trim();
    const rate  = REF && REF.toNumber ? REF.toNumber(row[colRateD -1]) : Number(row[colRateD -1] || 0);
    const revs  = REF && REF.toNumber ? REF.toNumber(row[colRevsC -1]) : Number(row[colRevsC -1] || 0);
    const own   = String(row[colOwnCat-1] || '').trim(); // «Своя категория»

    // G — всегда «Артикул»
    displayG.push(art);
    ratingD.push(rate);
    countC.push(revs);

    // ==== СС с фолбэком «Симкарты» через REF.resolveCCForArticle ====
    const cc = (REF && REF.resolveCCForArticle)
      ? REF.resolveCCForArticle(plat, art, own, ssAJ)
      : 0;
    ssAA.push(cc > 0 ? cc : 'нет СС');

    // ==== M — только «Наличие» из той же карты «🍔 СС» ====
    const tovar = (REF && REF.toTovarFromArticle) ? REF.toTovarFromArticle(plat, art) : art;
    const rec   = ssAJ.get(tovar); // {cc,nal,vput,vpost}
    const nal   = rec ? Number(rec.nal || 0) : 0;
    flowM.push(nal > 0 ? nal : '');

    // ==== N,O,P — как были из «Физ. оборот» ====
    const key = (REF && REF.makeSSKey) ? REF.makeSSKey(cab, art) : (cab + '␟' + art);
    const ph = physMap.get(key);

    if (ph) {
      const eNum = Number(ph.remainENum) || 0;
      const gNum = Number(ph.speedNumG)  || 0;

      // N — запас (E/G)
      let nVal = '';
      if (eNum === 0) nVal = '';
      else if (gNum === 0) nVal = 'нп';
      else {
        const div = eNum / gNum;
        nVal = (div === 0) ? '' : div;
      }

      // O — "E (F)"
      const fNum = Number(ph.inSuppFNum) || 0;
      let oStr = '';
      if      (eNum === 0 && fNum === 0) oStr = '';
      else if (eNum === 0 && fNum > 0)   oStr = '0 (' + fNum + ')';
      else if (eNum > 0  && fNum === 0)  oStr = String(eNum);
      else                                oStr = String(eNum) + ' (' + fNum + ')';

      // P — скорость (display), если G==0 → ''
      const pDisp = (gNum === 0) ? '' : (ph.speedDispG || '');

      flowN.push(nVal);
      flowO.push(oStr);
      flowP.push(pDisp);
    } else {
      flowN.push('');
      flowO.push('');
      flowP.push('');
    }
  });

  return { displayG, ratingD, countC, ssAA, flowM, flowN, flowO, flowP };
}

function emptyCalcRows_() {
  return { displayG: [], ratingD: [], countC: [], ssAA: [], flowM: [], flowN: [], flowO: [], flowP: [] };
}


/********************* ЗАПИСЬ В ЛИСТ ****************************/

function writeColumnG_(sh, displayValuesG, rowsLen) {
  var out = new Array(rowsLen);
  for (var i = 0; i < rowsLen; i++) {
    var v = (i < displayValuesG.length) ? displayValuesG[i] : '';
    out[i] = [ (v == null) ? '' : v ];
  }
  var rng = sh.getRange(ROW_DATA, col_('G'), rowsLen, 1);
  rng.setValues(out);
  rng.setNumberFormat('General');
  rng.setHorizontalAlignment('left');
}

function setGHeader_(sh) {
  sh.getRange(3, col_('G')).setValue('Артикул');
}

// «Отзывы» J:K (J = рейтинг, K = кол-во; нули -> пусто)
function writeReviews_(sh, ratingD, countC, rowsLen) {
  var outJ = new Array(rowsLen);
  var outK = new Array(rowsLen);

  for (var i = 0; i < rowsLen; i++) {
    var rRaw = (i < ratingD.length ? ratingD[i] : '');
    var kRaw = (i < countC.length  ? countC[i]  : '');

    var rNum = (rRaw === '' || rRaw == null) ? '' : Number(rRaw);
    var kNum = (kRaw === '' || kRaw == null) ? '' : Number(kRaw);

    var jVal = (rNum === 0) ? '' : (isFinite(rNum) ? rNum : '');
    var kVal = (kNum === 0) ? '' : (isFinite(kNum) ? kNum : '');

    outJ[i] = [jVal];
    outK[i] = [kVal];
  }

  sh.getRange(ROW_DATA, col_('J'), rowsLen, 1).setValues(outJ).setNumberFormat('General');
  sh.getRange(ROW_DATA, col_('K'), rowsLen, 1).setValues(outK).setNumberFormat('General');
}

// «СС» AA
function writeSS_(sh, ssAA, rowsLen) {
  var aaArr = ssAA.slice();
  while (aaArr.length < rowsLen) aaArr.push('');
  var rngAA = sh.getRange(ROW_DATA, col_('AA'), rowsLen, 1);
  rngAA.setValues(aaArr.map(v => [v])).setNumberFormat('General');
}


/********************* ОФОРМЛЕНИЕ *******************************/

function clearPrimerAfterSizing_(sh, rowsLen) {
  var lastRow = Math.max(ROW_DATA + rowsLen - 1, MIN_LAST_ROW);

  var rngRight = sh.getRange(ROW_DATA, col_('G'), lastRow - ROW_DATA + 1, col_('AC') - col_('G') + 1);
  rngRight.breakApart();
  rngRight.clearContent();
  rngRight.setBackground(COLOR.white);
  rngRight.setBorder(true, true, true, true, true, true, COLOR.white, SpreadsheetApp.BorderStyle.SOLID);

  if (lastRow >= 25) {
    var rngLeft = sh.getRange(25, 1, lastRow - 25 + 1, 6);
    rngLeft.breakApart();
    rngLeft.clearContent();
    rngLeft.setBackground(COLOR.white);
    rngLeft.setBorder(true, true, true, true, true, true, COLOR.white, SpreadsheetApp.BorderStyle.SOLID);
  }
}

function applyWidths_(sh) {
  var setW = function(label, px){ sh.setColumnWidth(col_(label), px); };
  ensureColCapacityTo_(sh, 29);
  var wF  = sh.getColumnWidth(col_('F'));
  var wAC = Math.max(1, Math.floor(wF / 2));
  sh.setColumnWidth(col_('AC'), wAC);

  // Разделители
  DIVIDERS.forEach(function(c){ sh.setColumnWidth(c, WIDTHS.separators); });

  // Узкие 62px — включая M
  WIDTHS.narrow62.forEach(function(cc){ setW(cc, 62); });

  setW('R', WIDTHS.Q);
  setW('U', WIDTHS.T);
  setW('S', WIDTHS.R);

  WIDTHS.other85.forEach(function(cc){ setW(cc, 85); });
}

function applyDataFormatting_Only_(sh, rowsLen) {
  if (rowsLen <= 0) return;
  var lastCol = sh.getMaxColumns();
  var colG = col_('G');

  sh.getRange(ROW_DATA, colG, rowsLen, Math.max(1, lastCol - colG + 1))
    .setFontFamily(FONT.family)
    .setFontSize(FONT.data)
    .setFontWeight('normal')
    .setFontColor(COLOR.txt)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  sh.getRange(ROW_DATA, col_('G'), rowsLen, 1).setHorizontalAlignment('left');
  sh.getRange(ROW_DATA, col_('H'), rowsLen, 1).setHorizontalAlignment('center');
}

function applyDataBackgrounds_(sh, rowsLen) {
  if (rowsLen <= 0) return;
  if (PALETTE.introF) sh.getRange(ROW_DATA, col_('G'), rowsLen, 1).setBackground(PALETTE.introF);
  if (PALETTE.flow)   sh.getRange(ROW_DATA, col_('M'), rowsLen, 4).setBackground(PALETTE.flow);
  if (PALETTE.calcT)  sh.getRange(ROW_DATA, col_('U'), rowsLen, 1).setBackground(PALETTE.calcT);
  if (PALETTE.profit) sh.getRange(ROW_DATA, col_('R'), rowsLen, 2).setBackground(PALETTE.profit);
}

function applyDataGrid_(sh) {
  var SOLID = SpreadsheetApp.BorderStyle.SOLID;
  var lastRow = sh.getMaxRows();
  var rows = Math.max(0, lastRow - ROW_DATA + 1);
  if (rows <= 0) return;

  var blocks = [
    { c1:col_('G'),  c2:col_('H')  }, // Вводные
    { c1:col_('J'),  c2:col_('K')  }, // Отзывы
    { c1:col_('M'),  c2:col_('P')  }, // Физ. оборот
    { c1:col_('R'),  c2:col_('S')  }, // Прибыль
    { c1:col_('U'),  c2:col_('Y')  }, // Расчёт начисления
    { c1:col_('AA'), c2:col_('AB') }  // Внешние (AA — СС; AB — налог, не пишем)
  ];

  blocks.forEach(function(b){
    var rng = sh.getRange(ROW_DATA, b.c1, rows, b.c2 - b.c1 + 1);
    rng.setBorder(null, null, null, null, true, true, COLOR.inner, SOLID);
    rng.setBorder(true, true, true, true, null, null, COLOR.outer, SOLID);
  });

  // акцентная вертикаль справа от U
  sh.getRange(ROW_DATA, col_('U'), rows, 1)
    .setBorder(null, null, null, true, null, null, COLOR.outer, SOLID);

  // прибьём правую грань "Прибыли"
  sh.getRange(ROW_DATA, col_('S'), rows, 1)
    .setBorder(null, null, null, true, null, null, COLOR.outer, SOLID);
}

function applyNumberFormatsRUAB_(sh, rowsLen) {
  if (rowsLen <= 0) return;

  // H — данные: формат #,##0 и центрирование
  sh.getRange(ROW_DATA, col_('H'), rowsLen, 1)
    .setNumberFormat('#,##0')
    .setHorizontalAlignment('center');

  // R — Прибыль
  sh.getRange(ROW_DATA, col_('R'), rowsLen, 1)
    .setNumberFormat('#,##0');

  // S — процент (0%)
  sh.getRange(ROW_DATA, col_('S'), rowsLen, 1)
    .setNumberFormat('0%');

  // U:AB — блок расчёта и внешние (включая AA=СС, AB=Налог)
  var fromU = col_('U');
  var toAB  = col_('AB');
  sh.getRange(ROW_DATA, fromU, rowsLen, toAB - fromU + 1)
    .setNumberFormat('#,##0');
}


/********************* ИСТОЧНИК «ФИЗ. ОБОРОТ» *******************/

/** Для калькулятора: читаем A:G и возвращаем Map(key -> данные), где key = "Кабинет␟Артикул" */
function readPhysMapForCabinet_(physSheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(physSheetName);
  const map = new Map();
  if (!sh) return map;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 7) return map;

  // Берём и значения, и дисплей (для G — скорость с форматом скобок)
  const vals = sh.getRange(2, 1, lastRow - 1, 7).getValues();          // A:G raw
  const disp = sh.getRange(2, 1, lastRow - 1, 7).getDisplayValues();   // A:G display

  for (var i = 0; i < vals.length; i++) {
    const rowV = vals[i], rowD = disp[i];
    const cab = String(rowV[0] || '').trim();
    const art = String(rowV[1] || '').trim();
    if (!cab || !art) continue;

    const key = (REF && REF.makeSSKey) ? REF.makeSSKey(cab, art) : (cab + '␟' + art);

    const toNum = (REF && REF.toNumber) ? REF.toNumber : function (v){ var n=Number(String(v).replace(',','.')); return isFinite(n)?n:0; };

    const stockCNum  = toNum(rowD[2]); // C
    const inWayDNum  = toNum(rowD[3]); // D
    const remainENum = toNum(rowD[4]); // E
    const inSuppFNum = toNum(rowD[5]); // F

    const speedDispG = String(rowD[6] || ''); // G display
    const speedNumG  = toNum(rowD[6]);        // G numeric

    map.set(key, {
      stockCNum, inWayDNum, remainENum, inSuppFNum,
      speedDispG, speedNumG
    });
  }
  return map;
}


/********************* ХЕЛПЕРЫ И УТИЛИТЫ ************************/

/** Мягкий поиск колонки по названию: игнорирует теги "[ OZ ]"/"[ WB ]", регистр и пробелы */
function findHeaderIndexFlexible_(headerRowValues, names) {
  const norm = (s) => String(s||'')
    .replace(/\[[^\]]*\]/g,'') // вырезаем [ OZ ] / [ WB ] и т.п.
    .trim()
    .toLowerCase();
  const hdr = headerRowValues.map(norm);
  const candidates = (names||[]).map(norm);
  for (let i = 0; i < hdr.length; i++) {
    if (candidates.indexOf(hdr[i]) !== -1) return i+1; // 1-based
  }
  return 0;
}

function col_(a1) { var n=0; for (var i=0;i<a1.length;i++) n=n*26+(a1.charCodeAt(i)-64); return n; }
function ensureRowCapacityTo_(sh, targetLastRow) {
  var maxRows = sh.getMaxRows();
  if (maxRows < targetLastRow) sh.insertRowsAfter(maxRows, targetLastRow - maxRows);
}
function trimRowsAfter_(sh, rowsLen) {
  var keepLast = Math.max(ROW_DATA + rowsLen - 1, MIN_LAST_ROW);
  var maxRows = sh.getMaxRows();
  if (keepLast < maxRows) sh.deleteRows(keepLast + 1, maxRows - keepLast);
}
function ensureColCapacityTo_(sh, minCols) {
  var maxCols = sh.getMaxColumns();
  if (maxCols < minCols) sh.insertColumnsAfter(maxCols, minCols - maxCols);
}
function rangeIntersects_(r, targetRange) {
  var r1=r.getRow(), r2=r1+r.getNumRows()-1;
  var c1=r.getColumn(), c2=c1+r.getNumColumns()-1;
  var t1=targetRange.getRow(), t2=t1+targetRange.getNumRows()-1;
  var k1=targetRange.getColumn(), k2=k1+targetRange.getNumColumns()-1;
  return !(r2 < t1 || r1 > t2 || c2 < k1 || c1 > k2);
}
function autoWidthPlus_(sh, colIndex, paddingPx) {
  sh.autoResizeColumn(colIndex);
  var w = sh.getColumnWidth(colIndex);
  sh.setColumnWidth(colIndex, Math.max(1, w + (Number(paddingPx) || 0)));
}

function writeFlowBlock_(sh, arrM, arrN, arrO, arrP, rowsLen) {
  const padVals = (src) => {
    const out = new Array(rowsLen);
    for (var i = 0; i < rowsLen; i++) out[i] = [ (i < src.length) ? src[i] : '' ];
    return out;
  };

  // M — только «Наличие», обычные значения
  sh.getRange(ROW_DATA, col_('M'), rowsLen, 1)
    .setValues(padVals(arrM))
    .setNumberFormat('General')
    .setHorizontalAlignment('left');

  // N — число (E/G), формат "0"; "нп" — строкой
  sh.getRange(ROW_DATA, col_('N'), rowsLen, 1)
    .setValues(padVals(arrN))
    .setNumberFormat('0')
    .setHorizontalAlignment('center');

  // O — "E (F)"
  sh.getRange(ROW_DATA, col_('O'), rowsLen, 1)
    .setValues(padVals(arrO))
    .setNumberFormat('General')
    .setHorizontalAlignment('left');

  // P — скорость (как в источнике)
  sh.getRange(ROW_DATA, col_('P'), rowsLen, 1)
    .setValues(padVals(arrP))
    .setNumberFormat('General')
    .setHorizontalAlignment('center');

  // ⛔️ БЕЗ автоширины для M
}
