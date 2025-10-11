/***************************************************************
 * getFORM — DUAL (OZ/WB)
 * Универсальный запуск layout без кулдаунов:
 *   runLayoutImmediate(selectedCab?)
 *
 * Источники выбираются строго через REF.getCurrentPlatform():
 *   - [OZ]/[WB] Артикулы  (A:M, 13 колонок; M = «Своя категория»)
 *   - [OZ]/[WB] Физ. оборот (A:D)
 *
 * Контрол кабинета — ИМЕНОВАННЫЙ ДИАПАЗОН REF.NAMED.CAB_CTRL (= 'muff_cabs')
 * «⛓️ Параллель» рендерится инлайном вместе с калькулятором (A:E + M)
 ***************************************************************/

//////////////////// Константы ////////////////////
const SHEET_CALC   = REF.SHEETS.CALC;
const SHEET_PARAMS = REF.SHEETS.PARAMS;

// источники листов
const ARTS_OZ  = REF.SHEETS.ARTS_OZ;
const ARTS_WB  = REF.SHEETS.ARTS_WB;
const PHYS_OZ  = REF.SHEETS.FIZ_OZ;
const PHYS_WB  = REF.SHEETS.FIZ_WB;

// единый контрол выбора кабинета (именованный диапазон)
const CTRL_RANGE_A1 = REF.NAMED.CAB_CTRL;
const CAB_PLACEHOLDER = '<выберите кабинет>';

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
const PALETTE = { introF: '#efefef', flow: '#fff2cc', calcT: '#e9e2f8', profit: '#fce5cd' };

/* ===== Параллель: внешний вид ===== */
const SHEET_PAR = '⛓️ Параллель';
const PAR_SEP_WIDTH = 3;
const PAR_HEAD_BG   = '#efefef';
const PAR_HEAD_FG   = '#000000';
const PAR_FONT_FAM  = 'Roboto';
const PAR_FONT_SIZE = 10;

/* ===== Внутренний профайлинг ===== */
const TECH_LOG_SHEET = '🛠 Тех. лог';
const TECH_LOG_FLAG_A1 = 'E1'; // чекбокс-флаг
const DP = PropertiesService.getDocumentProperties();
const KEY_TLOG_PREV = 'techlog_prev_enabled';

/** читаем флаг из E1 */
function isTechLogEnabled_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(TECH_LOG_SHEET);
    if (!sh) return false;
    const raw = String(sh.getRange(TECH_LOG_FLAG_A1).getDisplayValue() || '').trim().toUpperCase();
    return raw === 'TRUE' || raw === '1' || raw === 'ON' || raw === 'ДА';
  } catch (_) { return false; }
}

/**
 * Если флаг только что включили → очищаем A2:D и ставим заголовок.
 * Запоминаем состояние в DocumentProperties.
 * Возвращает текущее состояние (true/false).
 */
function maybeResetTechLogOnEnable_() {
  const cur = isTechLogEnabled_();
  try {
    const prev = String(DP.getProperty(KEY_TLOG_PREV) || '');
    if (cur && prev !== '1') {
      const ss = SpreadsheetApp.getActive();
      const sh = ss.getSheetByName(TECH_LOG_SHEET) || ss.insertSheet(TECH_LOG_SHEET);
      // Заголовок
      sh.getRange(1,1,1,4).setValues([['Phase','Rel(ms)','Message','Extra(JSON)']]);
      // Очистить тело
      const last = sh.getLastRow();
      if (last > 1) sh.getRange(2,1,last-1,4).clearContent();
      DP.setProperty(KEY_TLOG_PREV, '1');
    } else if (!cur && prev !== '0') {
      DP.setProperty(KEY_TLOG_PREV, '0');
    }
  } catch (_) {}
  return cur;
}

function techLog_(phase, t0, message, extraObj) {
  if (!isTechLogEnabled_()) return; // глобальный выключатель
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(TECH_LOG_SHEET);
    if (!sh) sh = ss.insertSheet(TECH_LOG_SHEET);
    // гарантируем заголовок
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Phase','Rel(ms)','Message','Extra(JSON)']);
    } else {
      const hdr = sh.getRange(1,1,1,4).getValues()[0].join('|');
      if (hdr !== 'Phase|Rel(ms)|Message|Extra(JSON)') {
        sh.getRange(1,1,1,4).setValues([['Phase','Rel(ms)','Message','Extra(JSON)']]);
      }
    }
    const rel = Date.now() - t0;
    sh.appendRow([String(phase||''), rel, String(message||''), extraObj ? JSON.stringify(extraObj) : '']);
  } catch(_){}
}
function paramsLogShort_(label, cabinets, plat) {
  try { REF.logRun && REF.logRun(label, cabinets, plat); } catch(_){}
}

function platformUiLabel_(tag) {
  return tag === 'OZ' ? 'OZON' : tag === 'WB' ? 'WILDBERRIES' : '';
}

/********************* ПУБЛИЧНЫЕ ХЕНДЛЕРЫ ************************/

function runLayoutImmediate(selectedCab) {
  const T0 = Date.now();
  techLog_('BEGIN', T0, 'runLayoutImmediate');
  // ── ПРЕФЛАЙТ «Цены OZ/WB»: если обе свежие — продолжаем; иначе покажем окно и выходим
  try {
    if (typeof CALC_preflightOrShowDialog === 'function') {
      var okToProceed = CALC_preflightOrShowDialog(); // true → оба прайс-источника свежие
      if (!okToProceed) {
        techLog_('PRE', T0, 'dispatcher shown (prices stale)');
        return; // окно выполнит обновления и затем само вызовет runLayoutImmediate()
      }
    }
  } catch(_) { /* мягко игнорируем */ }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_CALC);
  if (!sh) throw new Error(`Лист "${SHEET_CALC}" не найден`);

  const ctrl = getCabCtrlRange_();
  const currentCab = String(selectedCab || (ctrl ? ctrl.getDisplayValue() : '') || '').trim();
  if (!currentCab || currentCab === CAB_PLACEHOLDER) {
    ss.toast('Не выбран кабинет', 'Внимание', 3);
    techLog_('END', T0, 'abort: no cabinet');
    return;
  }

  // Кэш-контекст, читаем один раз
  const plat  = REF.getCurrentPlatform();                 // 'OZ' | 'WB' | null
  techLog_('CTX', T0, 'platform resolved', {plat});
  const ssAJ  = REF.readSS_AJ_Map ? REF.readSS_AJ_Map() : new Map();
  techLog_('SSAJ', T0, 'readSS_AJ_Map done', {size: (typeof ssAJ.size==='number'? ssAJ.size : 'n/a')});
  const ctx   = { plat: (plat==='OZ'||plat==='WB'?plat:'OZ'), ssAJ };

  // На время рендера отключаем валидацию
  removeCabinetDropdown_(ctrl);
  try {
    // Калькулятор
    techLog_('CALC_START', T0, 'layoutCalculator');
    try { layoutCalculator(currentCab, ctx); }
    catch(eCalc) { techLog_('CALC_ERR', T0, 'layoutCalculator error', {err:String(eCalc && eCalc.message || eCalc)}); throw eCalc; }
    techLog_('CALC_END', T0, 'layoutCalculator');

    // Параллель (инлайн)
    techLog_('PAR_CALL', T0, 'layoutParallelInline_');
    try { layoutParallelInline_(currentCab, ctx); }
    catch(ePar) { techLog_('PAR_ERR', T0, 'layoutParallelInline_ error', {err:String(ePar && ePar.message || ePar)}); throw ePar; }
    techLog_('PAR_CALL', T0, 'layoutParallelInline_');

    SpreadsheetApp.flush();
    techLog_('FLUSH', T0, 'SpreadsheetApp.flush()');

    paramsLogShort_('Калькулятор', currentCab, ctx.plat);
    techLog_('END', T0, 'All done');
  } finally {
    restoreCabinetDropdown_(ctrl, currentCab);
  }
}

/** onEdit: реагируем на muff_mp (площадка) и на muff_cabs (кабинет) */
function onEdit(e) {
  const T0 = Date.now();
  try {
    if (!e || !e.range) return;

    const ss  = SpreadsheetApp.getActive();
    const rng = e.range;
    const sh  = rng.getSheet();
    if (!sh) return;

    // ===== 1) Реакция на смену ПЛОЩАДКИ (именованный диапазон muff_mp) =====
    const rngPlat = safeGetRangeByName_(REF.NAMED.MP_CTRL);
    const hitPlat = rngPlat
      && rngPlat.getSheet().getSheetId() === sh.getSheetId()
      && rangeIntersects_(rng, rngPlat);

    if (hitPlat) {
      const raw = String(rngPlat.getDisplayValue() || '').trim();
      const tag = REF.platformCanon(raw) || 'OZ';

      // 1) Пересобрать список кабинетов под платформу и выставить плейсхолдер
      applyCabinetDropdownForCurrentPlatform_NoAutoSelect_();

      // 2) Рамка у кнопки-диапазона calc_button_refresh — под платформу OZ/WB
      try { applyCalcRefreshBorderForPlatform_(); } catch (_){}

      // 3) «Погасить» данные на Калькуляторе: шрифт = фон (эффект «пусто»)
      try { dimCalcDataFontsToBackground_(); } catch (_){}

      // 4) Уведомление пользователю (полное имя платформы)
      try { ss.toast('Площадка: ' + platformUiLabel_(tag) + '. Выберите кабинет.', 'Готово', 3); } catch (_){}

      // по ТЗ при смене площадки — НИЧЕГО не считаем
      return;
    }

    // ===== 2) Реакция на выбор КАБИНЕТА (именованный диапазон muff_cabs) =====
    if (sh.getName() !== SHEET_CALC) return;

    const ctrl = getCabCtrlRange_();
    if (!ctrl || !rangeIntersects_(rng, ctrl)) return;

    const selectedCab = String(ctrl.getDisplayValue() || '').trim();
    if (!selectedCab || selectedCab === CAB_PLACEHOLDER) return; // плейсхолдер/пусто → ничего не делаем

    const logEnabled = maybeResetTechLogOnEnable_();
    if (logEnabled) techLog_('invoke', T0, 'runLayoutImmediate');

    runLayoutImmediate(selectedCab);

    if (logEnabled) techLog_('DONE', T0, 'OK');
  } catch (err) {
    if (isTechLogEnabled_()) techLog_('ERROR', T0, String(err && err.message || err));
    throw err;
  }
}


/** Рамка вокруг именованного диапазона calc_button_refresh — под платформу OZ/WB (merge-aware, без Advanced API) */
function applyCalcRefreshBorderForPlatform_() {
  var plat = REF.getCurrentPlatform();
  var colorHex = (plat === 'WB') ? '#8c44bb' : (plat === 'OZ') ? '#016bbf' : null;
  if (!colorHex) return;

  var base = safeGetRangeByName_('calc_button_refresh');
  if (!base) return;

  // если именованный диапазон — часть мерджа, расширяем до всего мерджа
  var rng = base.isPartOfMerge() ? base.getMergedRanges()[0] : base;

  // Толстая внешняя рамка нужного цвета
  rng.setBorder(true, true, true, true, false, false, colorHex, SpreadsheetApp.BorderStyle.SOLID_THICK);
  SpreadsheetApp.flush();
}


function hexToRgbObj_(hex) {
  var h = String(hex || '').replace('#', '');
  if (h.length === 3) h = h.split('').map(function (x) { return x + x; }).join('');
  var r = parseInt(h.substr(0, 2), 16) / 255;
  var g = parseInt(h.substr(2, 2), 16) / 255;
  var b = parseInt(h.substr(4, 2), 16) / 255;
  return { red: r, green: g, blue: b };
}

/** «Погасить» данные на Калькуляторе: цвет шрифта = цвет фона для G:AC, строки данных */
function dimCalcDataFontsToBackground_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_CALC);
  if (!sh) return;

  var top = ROW_DATA;
  var last = Math.max(sh.getLastRow(), MIN_LAST_ROW);
  var rows = last - top + 1;
  if (rows <= 0) return;

  var c1 = col_('G'), c2 = col_('AC');
  var rng = sh.getRange(top, c1, rows, c2 - c1 + 1);
  var bgs = rng.getBackgrounds();      // 2D массив HEX фонов
  rng.setFontColors(bgs);              // ставим шрифт = фону (делаем «невидимым»)
}

/********************* КОНТРОЛ КАБИНЕТА **************************/

function setupCabinetControl_() {
  const ss = SpreadsheetApp.getActive();
  const shCalc = ss.getSheetByName(SHEET_CALC);
  if (!shCalc) throw new Error(`Лист "${SHEET_CALC}" не найден`);

  const ctrl = getCabCtrlRange_();
  const currentValue = String(ctrl ? (ctrl.getDisplayValue() || '') : '').trim();

  if (ctrl) {
    ctrl.breakApart();
    ctrl.clearDataValidations();
    ctrl.merge();
    ctrl.setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setFontFamily(FONT.family)
        .setFontSize(11);
  }

  restoreCabinetDropdown_(ctrl, currentValue || null);
}

function getCabCtrlRange_() {
  const ss = SpreadsheetApp.getActive();
  try {
    const r = REF.getCabinetControlRange && REF.getCabinetControlRange();
    if (r) return r;
  } catch(_) {}
  try { return ss.getRangeByName(CTRL_RANGE_A1); } catch(_) {}
  throw new Error('Именованный диапазон muff_cabs не найден');
}

function removeCabinetDropdown_(ctrlRange) { if (ctrlRange) ctrlRange.clearDataValidations(); }

function restoreCabinetDropdown_(ctrlRange, selectedCab) {
  const list = getCabinetListFromParams_(); // фильтруем по текущей платформе
  if (ctrlRange) ctrlRange.clearDataValidations();

  if (!list.length) { if (ctrlRange && selectedCab) ctrlRange.setValue(selectedCab); return; }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  if (ctrlRange) ctrlRange.setDataValidation(rule);

  if (!ctrlRange) return;
  const cur = String(ctrlRange.getDisplayValue() || '').trim();
  const chosen =
    (selectedCab && list.indexOf(selectedCab) !== -1) ? selectedCab :
    (list.indexOf(cur) !== -1) ? cur :
    list[0];

  ctrlRange.setValue(chosen);
}

/** Повесить список кабинетов по текущей платформе, но оставить плейсхолдер (не выбирать первый автоматически) */
function applyCabinetDropdownForCurrentPlatform_NoAutoSelect_() {
  const ctrl = getCabCtrlRange_();
  if (!ctrl) return;

  const list = getCabinetListFromParams_(); // фильтрует по REF.getCurrentPlatform()
  ctrl.clearDataValidations();

  if (list.length) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(list, true)
      .setAllowInvalid(false)
      .build();
    ctrl.setDataValidation(rule);
  }

  // ключевое: выставить плейсхолдер → не запускать расчёты
  ctrl.setValue(CAB_PLACEHOLDER);
}

/** Список кабинетов из «⚙️ Параметры», с учётом текущей платформы */
function getCabinetListFromParams_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_PARAMS);
  if (!sh) return [];

  const plat = REF.getCurrentPlatform(); // 'OZ' | 'WB' | null
  const last = sh.getLastRow();
  if (last < 2) return [];

  const rows = sh.getRange(2, 1, last - 1, 4).getValues(); // A..D (RAW!)
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const name = String(rows[i][0] || '').trim();               // A Кабинет
    const pRaw = String(rows[i][3] || '').trim().toUpperCase(); // D Площадка
    if (!name) continue;
    const p = REF.platformCanon(pRaw); // 'OZ' | 'WB' | null
    if (!plat || (p && p === plat)) out.push(name);
  }
  return Array.from(new Set(out));
}

/********************* ОСНОВНОЙ LAYOUT (КАЛЬКУЛЯТОР) *************/

function layoutCalculator(cabinet, ctx) {
  const T0 = Date.now();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_CALC);
  if (!sh) throw new Error(`Лист "${SHEET_CALC}" не найден`);

  techLog_('CALC', T0, 'collectRowsForCalculator_ begin');
  const src = collectRowsForCalculator_(cabinet, ctx);
  techLog_('CALC', T0, 'collectRowsForCalculator_ end', {rows: src.displayG.length});

  const rowsLen = Math.max(src.displayG.length, MIN_DATA_ROWS);

  const needLast = Math.max(ROW_DATA + rowsLen - 1, MIN_LAST_ROW);
  ensureColCapacityTo_(sh, 29);
  ensureRowCapacityTo_(sh, needLast);
  trimRowsAfter_(sh, rowsLen);

  clearPrimerAfterSizing_(sh, rowsLen);

  // G + заголовок + автоширина
  writeColumnG_(sh, src.displayG, rowsLen);
  setGHeader_(sh);
  autoWidthPlus_(sh, col_('G'), 50);

  // «Отзывы» (J:K) и «СС» (AA)
  writeReviews_(sh, src.ratingD, src.countC, rowsLen);
  writeSS_(sh, src.ssAA, rowsLen);

  // Блок M:P — M = только «Наличие»; N = E/G|нп; O = только остаток; P = скорость
  writeFlowBlock_(sh, src.flowM, src.flowN, src.flowO, src.flowP, rowsLen);

  applyNumberFormatsRUAB_(sh, rowsLen);
  applyWidths_(sh);
  applyDataFormatting_Only_(sh, rowsLen);
  applyDataBackgrounds_(sh, rowsLen);
  applyDataGrid_(sh);

  sh.getRange(ROW_DATA, col_('U'), rowsLen, 1).setFontWeight('bold');

  // Выделяем проблемные значения
  emphasizeNPandNoCC_(sh, rowsLen);

  techLog_('CALC', T0, 'layoutCalculator painted', {rowsLen});
}

/** ПЛАТФОРМА → 'OZ' | 'WB' (дефолт 'OZ' при null) */
function resolvePlatformCurrent_() {
  const tag = REF.getCurrentPlatform();
  if (tag === 'WB' || tag === 'OZ') return tag;
  try { SpreadsheetApp.getActive().toast('Площадка не задана — выбран OZ по умолчанию', 'Внимание', 3); } catch (_) {}
  return 'OZ';
}

function collectRowsForCalculator_(cabinet, ctx) {
  const ss  = SpreadsheetApp.getActive();
  const selectedCab = String(cabinet||'').trim();
  if (!selectedCab || selectedCab === CAB_PLACEHOLDER) return emptyCalcRows_();

  const plat = (ctx && ctx.plat) || resolvePlatformCurrent_(); // 'OZ' | 'WB'
  const artsSheetName = (plat === 'WB') ? ARTS_WB : ARTS_OZ;
  const physSheetName = (plat === 'WB') ? PHYS_WB : PHYS_OZ;

  const shS = ss.getSheetByName(artsSheetName);
  if (!shS) return emptyCalcRows_();

  const lastRow = shS.getLastRow();
  const lastCol = shS.getLastColumn();
  if (lastRow < 2 || lastCol < 13) return emptyCalcRows_();

  // ----- индексы колонок -----
  const headers = shS.getRange(1,1,1,13).getValues()[0];
  const colCab    = findHeaderIndexFlexible_(headers, ['Кабинет'])        || 1;   // A
  const colArt    = findHeaderIndexFlexible_(headers, ['Артикул'])        || 2;   // B
  const colRevsC  = findHeaderIndexFlexible_(headers, ['Отзывы'])         || 3;   // C
  const colRateD  = findHeaderIndexFlexible_(headers, ['Рейтинг'])        || 4;   // D
  const colFBO    = findHeaderIndexFlexible_(headers, ['FBO'])            || 6;   // F
  const colFBS    = findHeaderIndexFlexible_(headers, ['FBS'])            || 7;   // G
  const colVolI   = findHeaderIndexFlexible_(headers, ['Объем','Объём'])  || 9;   // I
  const colPriceJ = findHeaderIndexFlexible_(headers, ['Цена'])           || 10;  // J
  const colOwnCat = findHeaderIndexFlexible_(headers, ['Своя категория']) || 13;  // M

  const vals = shS.getRange(2,1,lastRow-1,13).getValues(); // RAW

  // ----- фильтр по кабинету -----
  const filtered = [];
  for (let i=0;i<vals.length;i++){
    const row = vals[i];
    const cab = String(row[colCab-1]||'').trim();
    const art = String(row[colArt-1]||'').trim();
    if (art && cab === selectedCab) filtered.push(row);
  }

  // стабильная сортировка по артикулу
  filtered.sort((a,b) => {
    const A = String(a[colArt-1]||'').trim();
    const B = String(b[colArt-1]||'').trim();
    return A < B ? -1 : (A > B ? 1 : 0);
  });

  const ssAJ = (ctx && ctx.ssAJ) ? ctx.ssAJ : (REF.readSS_AJ_Map ? REF.readSS_AJ_Map() : new Map());
  const physMap = readPhysMapForCabinet_(physSheetName); // A:B:C:D (C=Остаток, D=Скорость)

  // ----- данные для КАЛЬКУЛЯТОРА -----
  const displayG = [];
  const ratingD  = [];
  const countC   = [];
  const ssAA     = [];

  const flowM = [];
  const flowN = [];
  const flowO = [];
  const flowP = [];

  // ----- данные для ПАРАЛЛЕЛИ (те же строки/порядок) -----
  const parA = []; // Артикул
  const parB = []; // Цена
  const parC = []; // Объём
  const parD = []; // FBO
  const parE = []; // FBS
  // СС возьмём из ssAA (общий источник)

  for (let i=0;i<filtered.length;i++){
    const row  = filtered[i];
    const art  = String(row[colArt   -1] || '').trim();
    const rate = REF.toNumber(row[colRateD -1]);
    const revs = REF.toNumber(row[colRevsC -1]);
    const own  = String(row[colOwnCat-1] || '').trim();

    // суммарно для КАЛЬКУЛЯТОРА
    displayG.push(art);
    ratingD.push(rate);
    countC.push(revs);

    const cc = REF.resolveCCForArticle ? REF.resolveCCForArticle(plat, art, own, ssAJ) : 0;
    ssAA.push(cc > 0 ? cc : 'нет СС');

    const tovar = REF.toTovarFromArticle ? REF.toTovarFromArticle(plat, art) : art;
    const rec   = ssAJ.get(tovar);
    const nal   = rec ? Number(rec.nal || 0) : 0;
    flowM.push(nal > 0 ? nal : '');

    const key = REF.makeSSKey(selectedCab, art);
    const ph = physMap.get(key);
    if (ph) {
      const eNum = Number(ph.remainENum) || 0; // C — Остаток
      const gNum = Number(ph.speedNumG)  || 0; // D — Скорость (как число)

      // N — если остаток 0 → пусто; если скорость 0 → "нп"; иначе E/G
      let nVal = '';
      if (eNum === 0) nVal = '';
      else if (gNum === 0) nVal = 'нп';
      else nVal = (eNum / gNum) || '';

      // O — только Остаток (без «В поставке»)
      const oStr = (eNum === 0) ? '' : String(eNum);

      // P — скорость из D (display-значение, как на листе источника)
      const pDisp = ph.speedDispG || '';

      flowN.push(nVal);
      flowO.push(oStr);
      flowP.push(pDisp);
    } else {
      flowN.push('');
      flowO.push('');
      flowP.push('');
    }

    // ---- ПАРАЛЛЕЛЬ ----
    parA.push([art]);
    parB.push([row[colPriceJ-1]]);
    parC.push([row[colVolI  -1]]);
    parD.push([row[colFBO   -1]]);
    parE.push([row[colFBS   -1]]);
  }

  // Пакет для Параллели в ctx (общий кэш для этого прогона)
  if (ctx) {
    ctx.parallelCache = {
      cabinet: selectedCab,
      plat,
      A: parA,
      B: parB,
      C: parC,
      D: parD,
      E: parE,
      M: ssAA.map(v => [v]) // СС: тот же порядок, что и A..E
    };
  }

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

function setGHeader_(sh) { sh.getRange(3, col_('G')).setValue('Артикул'); }

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

  sh.getRange(ROW_DATA, col_('U'), rows, 1)
    .setBorder(null, null, null, true, null, null, COLOR.outer, SOLID);

  sh.getRange(ROW_DATA, col_('S'), rows, 1)
    .setBorder(null, null, null, true, null, null, COLOR.outer, SOLID);
}

function applyNumberFormatsRUAB_(sh, rowsLen) {
  if (rowsLen <= 0) return;

  sh.getRange(ROW_DATA, col_('H'), rowsLen, 1)
    .setNumberFormat('#,##0')
    .setHorizontalAlignment('center');

  sh.getRange(ROW_DATA, col_('R'), rowsLen, 1).setNumberFormat('#,##0');
  sh.getRange(ROW_DATA, col_('S'), rowsLen, 1).setNumberFormat('0%');

  var fromU = col_('U');
  var toAB  = col_('AB');
  sh.getRange(ROW_DATA, fromU, rowsLen, toAB - fromU + 1)
    .setNumberFormat('#,##0');
}

/********************* ИСТОЧНИК «ФИЗ. ОБОРОТ» *******************/

function readPhysMapForCabinet_(physSheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(physSheetName);
  const map = new Map();
  if (!sh) return map;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 4) return map; // нужно минимум A:D

  // Берём и raw, и display — чтоб корректно достать число и отображаемую строку
  const vals = sh.getRange(2, 1, lastRow - 1, 4).getValues();        // A:D raw
  const disp = sh.getRange(2, 1, lastRow - 1, 4).getDisplayValues(); // A:D display

  const toNum = (REF && REF.toNumber) ? REF.toNumber : function (v) {
    const n = Number(String(v).replace(',', '.')); return isFinite(n) ? n : 0;
  };

  for (var i = 0; i < vals.length; i++) {
    const rowV = vals[i], rowD = disp[i];

    const cab = String(rowV[0] || '').trim(); // A — Кабинет
    const art = String(rowV[1] || '').trim(); // B — Артикул
    if (!cab || !art) continue;

    const key = REF.makeSSKey(cab, art);

    // C — Остаток
    const remainENum = toNum(rowD[2]);        // числом (из display, с запятыми и т.п.)

    // D — Скорость
    const speedDispG = String(rowD[3] || ''); // как отображается на листе
    const speedNumG  = toNum(rowD[3]);        // числом

    // «В поставке» не используется и не читается — устанавливаем 0
    const inSuppFNum = 0;

    map.set(key, { remainENum, inSuppFNum, speedDispG, speedNumG });
  }

  return map;
}

/********************* «ПАРАЛЛЕЛЬ» — ИНЛАЙН (минимум) *********************/
function layoutParallelInline_(cabinetFull, ctx) {
  const T0 = Date.now();
  techLog_('PAR_START', T0, 'layoutParallelInline_');

  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_PAR);
  if (!sh) sh = ss.insertSheet(SHEET_PAR);

  const cabinet = String(cabinetFull || '').trim();
  if (!cabinet || cabinet === CAB_PLACEHOLDER) throw new Error('Не выбран кабинет для «⛓️ Параллель»');

  const plat = (ctx && (ctx.plat === 'OZ' || ctx.plat === 'WB')) ? ctx.plat : resolvePlatformCurrent_();

  // === 1) Пытаемся взять ГОТОВЫЕ ДАННЫЕ из калькулятора ===
  const pc = ctx && ctx.parallelCache;
  if (pc && pc.cabinet === cabinet && pc.plat === plat) {
    const n = pc.A.length;
    techLog_('PAR', T0, 'use parallelCache', { n });

    ensureRowsExactlyStrict_(sh, 1 + n);
    ensureColCapacityTo_(sh, Math.max(13, sh.getMaxColumns()));

    if (n > 0) {
      sh.getRange(2,  1, n, 1).setValues(pc.A); // A2:A (Артикул)
      sh.getRange(2,  2, n, 1).setValues(pc.B); // B2:B (Цена)
      sh.getRange(2,  3, n, 1).setValues(pc.C); // C2:C (Объём)
      sh.getRange(2,  4, n, 1).setValues(pc.D); // D2:D (FBO)
      sh.getRange(2,  5, n, 1).setValues(pc.E); // E2:E (FBS)
      sh.getRange(2, 13, n, 1).setValues(pc.M); // M2:M (СС)
    }

    techLog_('PAR_END', T0, 'layoutParallelInline_', { wrote: n, source: 'cache' });
    return;
  }

  // === 2) Fallback: собрать напрямую из артикулов (на всякий случай) ===
  const artsSheetName = (plat === 'WB') ? ARTS_WB : ARTS_OZ;
  const shS = ss.getSheetByName(artsSheetName);

  techLog_('PAR', T0, 'build data start (fallback)', { sheet: artsSheetName });

  let A = [], B = [], C = [], D = [], E = [], M = [];
  if (shS) {
    const lastRow = shS.getLastRow();
    if (lastRow >= 2 && shS.getLastColumn() >= 13) {
      const hdr = shS.getRange(1,1,1,13).getValues()[0];
      const cCab = findHeaderIndexFlexible_(hdr, ['Кабинет'])        || 1;  // A
      const cArt = findHeaderIndexFlexible_(hdr, ['Артикул'])        || 2;  // B
      const cFBO = findHeaderIndexFlexible_(hdr, ['FBO'])            || 6;  // F
      const cFBS = findHeaderIndexFlexible_(hdr, ['FBS'])            || 7;  // G
      const cVol = findHeaderIndexFlexible_(hdr, ['Объем','Объём'])  || 9;  // I
      const cPr  = findHeaderIndexFlexible_(hdr, ['Цена'])           || 10; // J
      const cOwn = findHeaderIndexFlexible_(hdr, ['Своя категория']) || 13; // M

      const vals = shS.getRange(2,1,lastRow-1,13).getValues();

      const rows = [];
      for (let i=0;i<vals.length;i++){
        const r = vals[i];
        const cab = String(r[cCab-1]||'').trim();
        const art = String(r[cArt-1]||'').trim();
        if (art && cab === cabinet) rows.push(r);
      }

      rows.sort((a,b) => {
        const A = String(a[cArt-1]||'').trim();
        const B = String(b[cArt-1]||'').trim();
        return A < B ? -1 : (A > B ? 1 : 0);
      });

      const ssAJ = (ctx && ctx.ssAJ) ? ctx.ssAJ : (REF.readSS_AJ_Map ? REF.readSS_AJ_Map() : new Map());

      const n = rows.length;
      A = new Array(n); B = new Array(n); C = new Array(n); D = new Array(n); E = new Array(n); M = new Array(n);

      for (let i=0;i<n;i++){
        const r   = rows[i];
        const art = String(r[cArt-1]||'').trim();
        const own = String(r[cOwn-1]||'').trim();

        A[i] = [art];
        B[i] = [r[cPr -1]];
        C[i] = [r[cVol-1]];
        D[i] = [r[cFBO-1]];
        E[i] = [r[cFBS-1]];

        let cc = 0;
        try { cc = REF.resolveCCForArticle ? REF.resolveCCForArticle(plat, art, own, ssAJ) : 0; }
        catch(e){
          if (isTechLogEnabled_()) techLog_('PAR_WARN', T0, 'CC resolve failed (fallback)', { art, err:String(e && e.message || e) });
        }
        M[i] = [cc > 0 ? cc : 'нет СС'];

        if (i % 120 === 0) techLog_('PAR_PROGRESS', T0, 'fallback rows', { i, total:n });
      }

      techLog_('PAR', T0, 'build data end (fallback)', { n });
    } else {
      techLog_('PAR', T0, 'build data end (fallback)', { n: 0, reason: 'empty arts sheet' });
    }
  } else {
    techLog_('PAR', T0, 'build data end (fallback)', { n: 0, reason: 'arts sheet missing' });
  }

  const n = A.length;
  ensureRowsExactlyStrict_(sh, 1 + n);
  ensureColCapacityTo_(sh, Math.max(13, sh.getMaxColumns()));
  if (n > 0) {
    sh.getRange(2,  1, n, 1).setValues(A);
    sh.getRange(2,  2, n, 1).setValues(B);
    sh.getRange(2,  3, n, 1).setValues(C);
    sh.getRange(2,  4, n, 1).setValues(D);
    sh.getRange(2,  5, n, 1).setValues(E);
    sh.getRange(2, 13, n, 1).setValues(M);
  }

  techLog_('PAR_END', T0, 'layoutParallelInline_', { wrote: n, source: 'fallback' });
}

/************* Вспомогалки *************/

function ensureRowsExactlyStrict_(sh, needRows) {
  const cur = sh.getMaxRows();
  if (needRows <= 0) return;
  if (cur < needRows) sh.insertRowsAfter(cur, needRows - cur);
  else if (cur > needRows) sh.deleteRows(needRows + 1, cur - needRows);
}

function ensureColCapacityTo_(sh, minCols) {
  const cur = sh.getMaxColumns();
  if (cur < minCols) sh.insertColumnsAfter(cur, minCols - cur);
}

/************* Хелперы «Параллели» (инлайн) *************/

function setParallelHeaders_(sh) {
  const hdrAE = [[ 'Артикул', 'Цена', 'Объём', 'Ставка \nFBO', 'Ставка \nFBS' ]];
  sh.getRange(1, 1, 1, 5).setValues(hdrAE);
  sh.getRange(1, 1, 1, 5)
    .setBackground(PAR_HEAD_BG)
    .setFontColor(PAR_HEAD_FG)
    .setFontFamily(PAR_FONT_FAM)
    .setFontSize(PAR_FONT_SIZE)
    .setFontWeight('normal')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // F1 — разделитель
  sh.getRange(1, 6).setValue('').setBackground('#ffffff').clearFormat();

  // M1 — «СС»
  ensureColCapacityTo_(sh, 13);
  sh.getRange(1, 13).setValue('СС');
}

function clearParallelTargets_(sh) {
  const maxRows = sh.getMaxRows();
  if (maxRows < 1) return;

  // A:F — шапка и данные
  sh.getRange(1, 1, 1, 6).clear();
  if (maxRows > 1) sh.getRange(2, 1, maxRows - 1, 6).clear();

  // M — только данные (M2..)
  if (maxRows > 1) {
    const rowsToClear = Math.max(maxRows - 1, 0);
    if (rowsToClear > 0) sh.getRange(2, 13, rowsToClear, 1).clear();
  }
}

// Ограниченные по высоте стили
function styleSeparatorColumn_(sh, colIndex, rowsCount) {
  const SOLID = SpreadsheetApp.BorderStyle.SOLID;
  const rows = Math.max(1, Number(rowsCount) || 1);
  const rng = sh.getRange(1, colIndex, rows, 1);
  rng.setBackground('#ffffff');
  // внутренние белые
  rng.setBorder(null, null, null, null, true, true, '#ffffff', SOLID);
  // внешние чёрные
  rng.setBorder(true, true, true, true, null, null, '#000000', SOLID);
}

function paintRightEdge_(sh, colIndex, rowsCount) {
  const SOLID = SpreadsheetApp.BorderStyle.SOLID;
  const rows = Math.max(1, Number(rowsCount) || 1);
  sh.getRange(1, colIndex, rows, 1)
    .setBorder(null, null, null, true, null, null, '#000000', SOLID);
}

/********************* ХЕЛПЕРЫ И УТИЛИТЫ ************************/

function findHeaderIndexFlexible_(headerRowValues, names) {
  const norm = (s) => String(s || '').replace(/\[[^\]]*\]/g, '').trim().toLowerCase();
  const hdr = headerRowValues.map(norm);
  const candidates = (names || []).map(norm);
  for (let i = 0; i < hdr.length; i++) { if (candidates.indexOf(hdr[i]) !== -1) return i + 1; }
  return 0;
}

function col_(a1) { var n = 0; for (var i = 0; i < a1.length; i++) n = n * 26 + (a1.charCodeAt(i) - 64); return n; }
function ensureRowCapacityTo_(sh, targetLastRow) {
  var maxRows = sh.getMaxRows();
  if (maxRows < targetLastRow) sh.insertRowsAfter(maxRows, targetLastRow - maxRows);
}
function trimRowsAfter_(sh, rowsLen) {
  var keepLast = Math.max(ROW_DATA + rowsLen - 1, MIN_LAST_ROW);
  var maxRows = sh.getMaxRows();
  if (keepLast < maxRows) sh.deleteRows(keepLast + 1, maxRows - keepLast);
}
function rangeIntersects_(r, targetRange) {
  var r1 = r.getRow(), r2 = r1 + r.getNumRows() - 1;
  var c1 = r.getColumn(), c2 = c1 + r.getNumColumns() - 1;
  var t1 = targetRange.getRow(), t2 = t1 + targetRange.getNumRows() - 1;
  var k1 = targetRange.getColumn(), k2 = k1 + targetRange.getNumColumns() - 1;
  return !(r2 < t1 || r1 > t2 || c2 < k1 || c1 > k2);
}

/** безопасный геттер именованных диапазонов (не падает, если имени нет) */
function safeGetRangeByName_(name) {
  try { return SpreadsheetApp.getActive().getRangeByName(name); } catch (_) { return null; }
}

function autoWidthPlus_(sh, colIndex, paddingPx) {
  sh.autoResizeColumn(colIndex);
  var w = sh.getColumnWidth(colIndex);
  sh.setColumnWidth(colIndex, Math.max(1, w + (Number(paddingPx) || 0)));
}

function writeFlowBlock_(sh, arrM, arrN, arrO, arrP, rowsLen) {
  const padVals = (src) => {
    const out = new Array(rowsLen);
    for (var i = 0; i < rowsLen; i++) out[i] = [(i < src.length) ? src[i] : ''];
    return out;
  };

  // M — только «Наличие»
  sh.getRange(ROW_DATA, col_('M'), rowsLen, 1)
    .setValues(padVals(arrM))
    .setNumberFormat('General')
    .setHorizontalAlignment('left');

  // N — число (E/G) или "нп" строкой
  sh.getRange(ROW_DATA, col_('N'), rowsLen, 1)
    .setValues(padVals(arrN))
    .setNumberFormat('0')
    .setHorizontalAlignment('center');

  // O — только Остаток (без «В поставке»)
  sh.getRange(ROW_DATA, col_('O'), rowsLen, 1)
    .setValues(padVals(arrO))
    .setNumberFormat('General')
    .setHorizontalAlignment('left');

  // P — скорость (как в источнике)
  sh.getRange(ROW_DATA, col_('P'), rowsLen, 1)
    .setValues(padVals(arrP))
    .setNumberFormat('General')
    .setHorizontalAlignment('center');
}

/** Выделить проблемные значения:
 *  - N: 'нп'  → красный жирный
 *  - AA: 'нет СС' → красный жирный
 */
function emphasizeNPandNoCC_(sh, rowsLen) {
  if (!sh || rowsLen <= 0) return;

  // ===== N (Запас)
  var rngN = sh.getRange(ROW_DATA, col_('N'), rowsLen, 1);
  var dispN = rngN.getDisplayValues();
  var colorsN = new Array(rowsLen);
  var weightsN = new Array(rowsLen);
  for (var i = 0; i < rowsLen; i++) {
    var s = String(dispN[i][0] || '').trim().toLowerCase();
    if (s === 'нп') {
      colorsN[i] = ['#cc0000'];
      weightsN[i] = ['bold'];
    } else {
      colorsN[i] = [COLOR.txt];
      weightsN[i] = ['normal'];
    }
  }
  rngN.setFontColors(colorsN).setFontWeights(weightsN);

  // ===== AA (СС)
  var rngAA = sh.getRange(ROW_DATA, col_('AA'), rowsLen, 1);
  var dispAA = rngAA.getDisplayValues();
  var colorsAA = new Array(rowsLen);
  var weightsAA = new Array(rowsLen);
  for (var j = 0; j < rowsLen; j++) {
    var t = String(dispAA[j][0] || '').trim().toLowerCase();
    if (t === 'нет сс') {
      colorsAA[j] = ['#cc0000'];
      weightsAA[j] = ['bold'];
    } else {
      colorsAA[j] = [COLOR.txt];
      weightsAA[j] = ['normal'];
    }
  }
  rngAA.setFontColors(colorsAA).setFontWeights(weightsAA);
}

/************* Размерность для «Параллели» *************/
function ensureRowsExactly_(sh, needRows) {
  const cur = sh.getMaxRows();
  if (cur < needRows)      sh.insertRowsAfter(cur, needRows - cur);
  else if (cur > needRows) sh.deleteRows(needRows + 1, cur - needRows);
}

/** Имя листа артикулов по платформе */
function getArtsSheetNameByPlat_(plat) {
  return (plat === 'WB') ? ARTS_WB : ARTS_OZ;
}

/** Считает максимум позиций (артикулов) среди кабинетов на текущей площадке */
function computeMaxCabinetArticles_(plat) {
  const ss = SpreadsheetApp.getActive();
  const artsName = getArtsSheetNameByPlat_(plat);
  const sh = ss.getSheetByName(artsName);
  if (!sh) return 0;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return 0;

  // Читаем A:B (Кабинет, Артикул) разом
  const rng = sh.getRange(2, 1, lastRow - 1, Math.min(13, lastCol));
  const vals = rng.getValues();

  // Узнаём индексы колонок (на случай переименований)
  const hdr = sh.getRange(1,1,1,Math.min(13,lastCol)).getValues()[0];
  const cCab = findHeaderIndexFlexible_(hdr, ['Кабинет']) || 1; // A
  const cArt = findHeaderIndexFlexible_(hdr, ['Артикул']) || 2; // B

  const counts = new Map(); // cabinet -> count
  for (let i = 0; i < vals.length; i++) {
    const cab = String(vals[i][cCab-1] || '').trim();
    const art = String(vals[i][cArt-1] || '').trim();
    if (!cab || !art) continue;
    counts.set(cab, (counts.get(cab) || 0) + 1);
  }

  let maxCount = 0;
  counts.forEach(v => { if (v > maxCount) maxCount = v; });
  return maxCount;
}

/**
 * Убедиться, что на «⛓️ Параллель» строк ровно (1 шапка + maxArticles).
 * Меняем количество строк ТОЛЬКО если реально отличается (без «дыхания»).
 */
function ensureParallelRowsByMaxCabinet_(sh, plat) {
  const maxArticles = computeMaxCabinetArticles_(plat);
  const want = 1 + Math.max(0, maxArticles); // +1 — шапка
  const cur  = sh.getMaxRows();
  if (want <= 0) return;

  if (cur < want) {
    sh.insertRowsAfter(cur, want - cur);
  } else if (cur > want) {
    // удаляем только «хвост» одним вызовом
    sh.deleteRows(want + 1, cur - want);
  }
}
