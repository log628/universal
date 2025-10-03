/**
 * getParallel.gs — «⛓️ Параллель» (OZ/WB)
 * Пишем ТОЛЬКО A:F + M:
 *   A2:A — Артикул (из того же источника, что калькулятор; НЕ копируя из калькулятора)
 *   B:E  — Цена, Объём, Ставка FBO, Ставка FBS
 *   F    — узкий разделитель (фон белый; внутр. границы — белые; внешние — чёрные)
 *   M2:M — СС: через REF.resolveCCForArticle(platform, article, ownCategory, ssAJ)
 *           (по «🍔 СС»!A:J; фолбэк «Симкарты»: L="симка", M*2)
 * Перед вставкой чистим ТОЛЬКО A:F и M (данные; заголовок M1 не трогаем).
 * Высоту листа подгоняем под n данных.
 *
 * ВАЖНО: Площадка берётся строго из ⚙️ Параметры!I2 (OZON/WILDBERRIES).
 * Если I2 пусто, мягкий фолбэк — по наличию кабинета на листах артикулов.
 */

//////////////////// Константы ////////////////////

// Внешний вид «Параллели»
const PAR_SEP_WIDTH = 3;          // ширина узких разделителей (колонки F и P/16)
const PAR_HEAD_BG   = '#efefef';  // фон шапки A:E
const PAR_HEAD_FG   = '#000000';  // цвет текста шапки
const PAR_FONT_FAM  = 'Roboto';   // шрифт шапки
const PAR_FONT_SIZE = 10;         // размер шрифта шапки
const SHEET_PAR      = '⛓️ Параллель';     // локальное имя этого листа
const SHEET_CALC_P   = REF.SHEETS.CALC;    // калькулятор
const SHEET_PARAMS_P = REF.SHEETS.PARAMS;  // параметры

// Источники артикулов строго из REF
const ARTS_OZ_P = REF.SHEETS.ARTS_OZ;
const ARTS_WB_P = REF.SHEETS.ARTS_WB;

// Общий A1-диапазон контрола
const CTRL_RANGE_A1_P = REF.CTRL_RANGE_A1;

//////////////////// Публичная ////////////////////
function layoutParallel(cabinetOverride) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_PAR);
  if (!sh) sh = ss.insertSheet(SHEET_PAR);

  // 1) Текущий кабинет
  const cabinet = String(cabinetOverride || getCurrentCabinet_() || '').trim();
  if (!cabinet) throw new Error('Не выбран кабинет (⚖️ Калькулятор!B3:E4)');

  // 2) Источник и подготовка данных (включая CC)
  const src = buildRowsForParallel_(cabinet); // {display, price, volume, fbo, fbs, ss}
  const n = src.display.length;

  // 3) Высота листа: шапка (1) + минимум 5 строк данных даже при n=0
  const MIN_DATA_ROWS = 5;
  const dataRows = Math.max(n, MIN_DATA_ROWS);
  const needRows = 1 + dataRows;
  ensureRowsExactly_(sh, needRows);

  // 4) Очистка ТОЛЬКО A:F и M2..
  clearTargetColumns_(sh);

  // 5) Шапка
  setHeaders_(sh);

  // 6) Данные A2:F
  if (n > 0) {
    sh.getRange(2, 1, n, 1).setValues(src.display.map(v => [v])).setNumberFormat('General').setHorizontalAlignment('left'); // A
    sh.getRange(2, 2, n, 1).setValues(src.price  .map(v => [v])).setNumberFormat('General').setHorizontalAlignment('left'); // B
    sh.getRange(2, 3, n, 1).setValues(src.volume .map(v => [v])).setNumberFormat('General').setHorizontalAlignment('left'); // C
    sh.getRange(2, 4, n, 1).setValues(src.fbo    .map(v => [v])).setNumberFormat('General').setHorizontalAlignment('left'); // D
    sh.getRange(2, 5, n, 1).setValues(src.fbs    .map(v => [v])).setNumberFormat('General').setHorizontalAlignment('left'); // E
    sh.getRange(2, 6, n, 1).setNumberFormat('General'); // F пустая
  }

  // 7) Разделители: F и P
  ensureColCapacityTo_(sh, 16);
  styleSeparatorColumn_(sh, 6);
  styleSeparatorColumn_(sh, 16);
  sh.setColumnWidth(6,  PAR_SEP_WIDTH);
  sh.setColumnWidth(16, PAR_SEP_WIDTH);

  // 8) M2:M — СС
  ensureColCapacityTo_(sh, 13);
  if (n > 0) {
    sh.getRange(2, 13, n, 1)
      .setValues(src.ss.map(v => [v]))
      .setNumberFormat('General')
      .setHorizontalAlignment('left');
  }

  // 9) Чёрные правые грани G, M, U, W
  ensureColCapacityTo_(sh, 23);
  paintRightEdge_(sh, 7);
  paintRightEdge_(sh, 13);
  paintRightEdge_(sh, 21);
  paintRightEdge_(sh, 23);
}


//////////////////// Построение данных ////////////////////
function buildRowsForParallel_(cabinet) {
  const ss  = SpreadsheetApp.getActive();

  // Площадка: строго по I2; если пусто — фолбэки
  const plat = resolvePlatformForCabinet_PAR_(cabinet); // 'OZ' | 'WB'
  const artsSheetName = (plat === 'WB') ? ARTS_WB_P : ARTS_OZ_P;

  const shS = ss.getSheetByName(artsSheetName);
  if (!shS) return emptyPAR_();

  const lastRow = shS.getLastRow();
  const lastCol = shS.getLastColumn();
  // теперь ждём 13 колонок (A:M), где M = «Своя категория»
  if (lastRow < 2 || lastCol < 13) return emptyPAR_();

  const headers = shS.getRange(1,1,1,13).getDisplayValues()[0];

  const colCab    = findHeaderIndexFlexible_(headers, ['Кабинет'])         || 1;
  const colArt    = findHeaderIndexFlexible_(headers, ['Артикул'])         || 2;
  const colFBO    = findHeaderIndexFlexible_(headers, ['FBO'])             || 6;
  const colFBS    = findHeaderIndexFlexible_(headers, ['FBS'])             || 7;
  const colVol    = findHeaderIndexFlexible_(headers, ['Объем','Объём'])   || 9;
  const colPrice  = findHeaderIndexFlexible_(headers, ['Цена'])            || 10;
  const colOwnCat = findHeaderIndexFlexible_(headers, ['Своя категория'])  || 13; // NEW

  const vals = shS.getRange(2,1,lastRow-1,13).getDisplayValues();

  // Фильтр по кабинету
  const filtered = vals.filter(row => {
    const cab  = String(row[colCab -1] || '').trim();
    const art  = String(row[colArt -1] || '').trim();
    return art && cab === cabinet;
  });

  // Сортировка по Артикулу
  filtered.sort((a,b) => {
    const A = String(a[colArt-1]||'').trim();
    const B = String(b[colArt-1]||'').trim();
    return A.localeCompare(B, 'ru');
  });

  // Карта «🍔 СС» A:J (Товар -> {cc,nal,vput,vpost})
  const ssAJ = (typeof REF !== 'undefined' && REF.readSS_AJ_Map) ? REF.readSS_AJ_Map() : new Map();

  const display = []; // A — Артикул
  const price   = []; // B
  const volume  = []; // C
  const fbo     = []; // D
  const fbs     = []; // E
  const ssOut   = []; // M — СС

  filtered.forEach(row => {
    const art = String(row[colArt   -1] || '').trim();
    const own = String(row[colOwnCat-1] || '').trim(); // «Своя категория»

    display.push(art);
    price.push(row[colPrice-1]);
    volume.push(row[colVol  -1]);
    fbo.push(   row[colFBO -1]);
    fbs.push(   row[colFBS -1]);

    // СС через единый резолвер (с фолбэком «Симкарты»)
    let cc = 0;
    if (typeof REF !== 'undefined' && REF.resolveCCForArticle) {
      cc = REF.resolveCCForArticle(plat, art, own, ssAJ);
    } else {
      // на всякий случай старый путь (если нет резолвера)
      const tovar = (REF && REF.toTovarFromArticle) ? REF.toTovarFromArticle(plat, art) : art;
      const rec = ssAJ.get(tovar);
      cc = (rec && isFinite(Number(rec.cc)) && Number(rec.cc) > 0) ? Number(rec.cc) : 0;
    }
    ssOut.push(cc > 0 ? cc : 'нет СС');
  });

  return { display, price, volume, fbo, fbs, ss: ssOut };
}




function emptyPAR_() { return { display:[], price:[], volume:[], fbo:[], fbs:[], ss:[] }; }

//////////////////// Заголовки ////////////////////
function setHeaders_(sh) {
  // A:E — Артикул, Цена, Объём, "Ставка FBO", "Ставка FBS"
  const hdrAE  = [[ 'Артикул', 'Цена', 'Объём', 'Ставка \nFBO', 'Ставка \nFBS' ]];
  sh.getRange(1,1,1,5).setValues(hdrAE);
  sh.getRange(1,1,1,5)
    .setBackground(PAR_HEAD_BG)
    .setFontColor(PAR_HEAD_FG)
    .setFontFamily(PAR_FONT_FAM)
    .setFontSize(PAR_FONT_SIZE)
    .setFontWeight('normal')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // F1 — пусто и белое (узкий разделитель)
  sh.getRange(1,6).setValue('').setBackground('#ffffff').clearFormat();

  // M1 — «СС»
  ensureColCapacityTo_(sh, 13);
  sh.getRange(1,13).setValue('СС');
}




//////////////////// Очистка целевых колонок ////////////////////
function clearTargetColumns_(sh) {
  const maxRows = sh.getMaxRows();
  if (maxRows < 1) return;

  // A:F — шапка и данные
  sh.getRange(1,1,1,6).clear();
  if (maxRows > 1) sh.getRange(2,1,maxRows-1,6).clear();

  // M — только данные (M2..)
  if (maxRows > 1) {
    const rowsToClear = Math.max(maxRows - 1, 0);
    if (rowsToClear > 0) sh.getRange(2,13,rowsToClear,1).clear();
  }
}

//////////////////// Стили разделителей и граней ////////////////////
function styleSeparatorColumn_(sh, colIndex) {
  const SOLID = SpreadsheetApp.BorderStyle.SOLID;
  const maxRows = sh.getMaxRows();
  if (maxRows < 1) return;

  const rng = sh.getRange(1, colIndex, maxRows, 1);
  rng.setBackground('#ffffff');                                        // фон колонки — белый
  rng.setBorder(null, null, null, null, true, true, '#ffffff', SOLID); // внутренние — белые
  rng.setBorder(true,  true,  true,  true,  null, null, '#000000', SOLID); // внешние — чёрные
}

function paintRightEdge_(sh, colIndex) {
  const SOLID = SpreadsheetApp.BorderStyle.SOLID;
  const maxRows = sh.getMaxRows();
  if (maxRows < 1) return;
  sh.getRange(1, colIndex, maxRows, 1)
    .setBorder(null, null, null, true, null, null, '#000000', SOLID);
}

//////////////////// Резолвер площадки ////////////////////
function resolvePlatformForCabinet_PAR_(cabinet) {
  const ss = SpreadsheetApp.getActive();
  const shParams = ss.getSheetByName(SHEET_PARAMS_P);
  const filterUP = (shParams ? String(shParams.getRange('I2').getDisplayValue() || '').trim().toUpperCase() : '');

  if (filterUP === 'OZON' || filterUP === 'OZ') return 'OZ';
  if (filterUP === 'WILDBERRIES' || filterUP === 'WB') return 'WB';

  // Если фильтр не задан — попробуем найти кабинет в листах артикулов
  const foundOZ = cabinetExistsOnSheet_(ARTS_OZ_P, cabinet);
  const foundWB = cabinetExistsOnSheet_(ARTS_WB_P, cabinet);

  if (foundOZ && !foundWB) return 'OZ';
  if (!foundOZ && foundWB) return 'WB';

  // Дефолт
  return 'OZ';
}

function cabinetExistsOnSheet_(sheetName, cabinet) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return false;
  const last = sh.getLastRow();
  if (last < 2) return false;
  const vals = sh.getRange(2, 1, last - 1, 1).getDisplayValues(); // A = Кабинет
  const target = String(cabinet || '').trim();
  for (let i=0;i<vals.length;i++){
    if (String(vals[i][0] || '').trim() === target) return true;
  }
  return false;
}

//////////////////// Мягкий поиск заголовков ////////////////////
function findHeaderIndexFlexible_(headerRowValues, names) {
  const norm = (s) => String(s||'')
    .replace(/\[[^\]]*\]/g,'') // вырезаем [ OZ ] / [ WB ]
    .trim()
    .toLowerCase();
  const hdr = headerRowValues.map(norm);
  const candidates = (names||[]).map(norm);
  for (let i = 0; i < hdr.length; i++) {
    if (candidates.indexOf(hdr[i]) !== -1) return i+1; // 1-based
  }
  return 0;
}

//////////////////// Утилиты ////////////////////
function ensureRowsExactly_(sh, needRows) {
  const cur = sh.getMaxRows();
  if (cur < needRows)      sh.insertRowsAfter(cur, needRows - cur);
  else if (cur > needRows) sh.deleteRows(needRows + 1, cur - needRows);
}
function ensureColCapacityTo_(sh, minCols) {
  const cur = sh.getMaxColumns();
  if (cur < minCols) sh.insertColumnsAfter(cur, minCols - cur);
}
function getCurrentCabinet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_CALC_P);
  if (!sh) return '';
  return String(sh.getRange(CTRL_RANGE_A1_P).getDisplayValue() || '').trim();
}


