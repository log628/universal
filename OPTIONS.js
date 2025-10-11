/* ======================  KBR_ARROWS (изолированный модуль)  ======================
 * Публичные функции для рисунков-кнопок:
 *   toggleArrow_method()
 *   toggleArrow_raschet()
 *   toggleArrow_procent()
 *   toggleArrow_enterprice()
 *
 * Источник истины: лист «⚙️ Параметры», колонки K:N.
 *  - K: настройка ('метод' | 'расчет' | 'процент' | 'ввод цены' — нижний регистр)
 *  - L: «Включение» — текущее значение (равно M → левая; равно N → правая)
 *  - M: «Опция 1» — левая сторона
 *  - N: «Опция 2» — правая сторона
 *
 * Цвет текста не меняем. Выравнивание в B/E не трогаем.
 * Независимость: по умолчанию перерисовывается только нажатая кнопка (RESTYLE_ALL_AFTER_FLIP=false).
 */

var KBR_ARROWS = KBR_ARROWS || (function () {
  // ===== Настройки поведения
  const RESTYLE_ALL_AFTER_FLIP = false; // если true — перерисуем все 3 кнопки после клика
  const INVERT_FLOATING_BORDER = true;  // если грань выглядит «зеркально», поставьте true

  // ===== Листы
const KBR_SHEET_CALC_NAME   = (REF && REF.SHEETS && REF.SHEETS.CALC)   || '⚖️ Калькулятор';
const KBR_SHEET_PARAMS_NAME = (REF && REF.SHEETS && REF.SHEETS.PARAMS) || '⚙️ Параметры';

  // ===== Ключи настроек и подписи
  const KBR_BTN_KEYS     = ['метод','расчет','процент'];
  const KBR_BTN_LABELSUI = ['Метод','Расчёт','Процент'];

  // Ключ для быстрого заполнения
  const KBR_ENTERPRICE_LABEL_KEY = 'ввод цены';
  const KBR_ENTERPRICE_LABEL_UI  = 'Ввод цены';

  // ===== Палитра (цвет текста не используем)
  const BG_ACTIVE     = '#efefef';
  const BG_INACTIVE   = '#999999';

  // Для updateBorders
  const KBR_COLOR_BLACK  = { red: 0, green: 0, blue: 0 };
  const KBR_COLOR_EFEFEF = { red: 239/255, green: 239/255, blue: 239/255 };

  // Размер шрифта
  const FONT_ACTIVE_SIZE   = 10;
  const FONT_INACTIVE_SIZE = 9;

  // --------------------------------------------------------------------------------
  //                                  ПУБЛИЧНЫЕ
  // --------------------------------------------------------------------------------

  // Переключение: метод/расчёт/процент — строго от «⚙️ Параметры»
  function flip(idx) {
    const key   = KBR_BTN_KEYS[idx];
    const label = KBR_BTN_LABELSUI[idx];

    const ss = SpreadsheetApp.getActive();
    const shCalc = ss.getSheetByName(KBR_SHEET_CALC_NAME);
    if (!shCalc) { ss.toast('⚠️ Лист «⚖️ Калькулятор» не найден', 'Ошибка', 4); return; }

    const curSide  = getSideFromParams_(key);                 // 'left' | 'right'
    const nextSide = (curSide === 'left') ? 'right' : 'left';
    const chosenValue = setSideInParams_(key, nextSide);      // L = M|N

    const settingsTopRow = findHeaderTopRow_(shCalc, 'Настройки');
    const btnRange = getSettingsButtonRange_(shCalc, settingsTopRow, idx); // C:D (2 строки)
    if (!btnRange) { ss.toast('⚠️ Не найден блок «Настройки»', 'Ошибка', 4); return; }

    // Рамки кнопки (левая/правая грань) — по стороне
    paintBordersBySide_(btnRange, nextSide);

    // Подписи B/E — по стороне (фон, жирность, размер). ВЫРАВНИВАНИЕ НЕ МЕНЯЕМ.
    const rowTop = btnRange.getRow();
    const rngB = mergeAware_(shCalc.getRange(rowTop, 2, 2, 1));
    const rngE = mergeAware_(shCalc.getRange(rowTop, 5, 2, 1));
    if (nextSide === 'left') {
      applySideStyle_(rngB, true);
      applySideStyle_(rngE, false);
    } else {
      applySideStyle_(rngB, false);
      applySideStyle_(rngE, true);
    }

    if (RESTYLE_ALL_AFTER_FLIP) {
      restyleAllOptions_(shCalc, settingsTopRow);
    }

    ss.toast('📃 ' + label + ' = ' + (chosenValue || '—'), 'Готово', 3);
  }

  // Быстрое заполнение «Ввод цены»: очистка/заполнение H4:H{по G}
  // Кнопку «Быстрое заполнение» не подкрашиваем и выравнивание не меняем.
  function flipEnterPrice() {
    const ss = SpreadsheetApp.getActive();
    const shCalc = ss.getSheetByName(KBR_SHEET_CALC_NAME);
    const shPar  = ss.getSheetByName(KBR_SHEET_PARAMS_NAME);
    if (!shCalc || !shPar) return;

    // Прочитать L("ввод цены") и L("процент")
    const lastPar = Math.max(2, shPar.getLastRow());
    const rowsPar = (lastPar >= 2) ? shPar.getRange(2, 11, lastPar - 1, 4).getDisplayValues() : []; // K..N
    let stateEnterPrice = ''; // L по "ввод цены"
    let statePercent    = ''; // L по "процент"
    for (let i = 0; i < rowsPar.length; i++) {
      const key = String(rowsPar[i][0] || '').trim().toLowerCase(); // K
      const val = String(rowsPar[i][1] || '').trim();               // L
      if (key === 'ввод цены') stateEnterPrice = val;
      if (key === 'процент')   statePercent    = val;
    }

    // Рабочая высота по последней непустой в G4:G
    const lastRowCalc = shCalc.getLastRow();
    const height      = Math.max(lastRowCalc - 3, 1);
    const gVals       = shCalc.getRange(4, 7, height, 1).getValues(); // G4:G
    let lastIdx = -1;
    for (let i = gVals.length - 1; i >= 0; i--) {
      if (String(gVals[i][0]).trim() !== '') { lastIdx = i; break; }
    }
    if (lastIdx < 0) { ss.toast('📝 Ввод цены: заполнено', 'Готово', 3); return; }

    const endRow = 4 + lastIdx;
    theNum = endRow - 3; // keep variable for next line
    const num    = theNum;
    const hRange = shCalc.getRange(4, 8, num, 1); // H4:H{endRow}

    if (stateEnterPrice === 'Заполнено') {
      // ОЧИСТКА
      hRange.clearContent();
      hRange.setNumberFormat('#,##0');
      ss.toast('🧹 Ввод цены: очищено', 'Готово', 3);
    } else {
      // ЗАПОЛНЕНИЕ
      const shParallel = ss.getSheetByName('⛓️ Параллель');
      if (!shParallel || shParallel.getLastRow() < 2) {
        ss.toast('📝 Ввод цены: заполнено', 'Готово', 3);
        return;
      }

      const srcColIndex = (statePercent === 'Фактич. %') ? 17 : 2; // Q=17, B=2
      const lastParRow = shParallel.getLastRow();

      const parA = shParallel.getRange(2, 1,  lastParRow - 1, 1).getValues();           // A2:A
      const parM = shParallel.getRange(2, 13, lastParRow - 1, 1).getDisplayValues();    // M2:M
      const parV = shParallel.getRange(2, srcColIndex, lastParRow - 1, 1).getValues();  // B2:B или Q2:Q

      const map = new Map(); // art -> value (только если M != "нет СС")
      for (let i = 0; i < parA.length; i++) {
        const art = String(parA[i][0] || '').trim();
        if (!art) continue;
        const mark = String(parM[i][0] || '').trim().toLowerCase();
        if (mark === 'нет сс') continue;
        map.set(art, parV[i][0]);
      }

      const out = new Array(num);
      const cur = hRange.getValues();
      for (let i = 0; i < num; i++) {
        const art = String(gVals[i][0] || '').trim();
        if (map.has(art)) {
          let v = map.get(art);
          const n = (typeof v === 'number') ? v : Number(String(v).replace(',', '.'));
          out[i] = [Number.isFinite(n) ? n : ''];
        } else {
          out[i] = [cur[i][0]];
        }
      }

      hRange.setValues(out);
      hRange.setNumberFormat('#,##0');
      ss.toast('📝 Ввод цены: заполнено', 'Готово', 3);
    }
  }

  // --------------------------------------------------------------------------------
  //                               ВСПОМОГАТЕЛЬНЫЕ
  // --------------------------------------------------------------------------------

  // Читаем сторону из Параметров: 'left' если L==M, 'right' если L==N, иначе 'right'
  function getSideFromParams_(settingKey) {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(KBR_SHEET_PARAMS_NAME);
    if (!sh) return 'right';

    const last = Math.max(2, sh.getLastRow());
    if (last < 2) return 'right';

    const rows = sh.getRange(2, 11, last - 1, 4).getDisplayValues(); // K..N
    const want = String(settingKey || '').trim().toLowerCase();

    for (let i = 0; i < rows.length; i++) {
      const k = String(rows[i][0] || '').trim().toLowerCase(); // K
      if (k !== want) continue;
      const L = String(rows[i][1] || '').trim(); // L
      const M = String(rows[i][2] || '').trim(); // M
      const N = String(rows[i][3] || '').trim(); // N
      if (L && L === M) return 'left';
      if (L && L === N) return 'right';
      break;
    }
    return 'right';
  }

  // Записываем сторону в L (L = M|N). Создаём лист/шапку/строку при необходимости. Возвращаем значение для тоста.
  function setSideInParams_(settingKey, side /* 'left'|'right' */) {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(KBR_SHEET_PARAMS_NAME) || ss.insertSheet(KBR_SHEET_PARAMS_NAME);

    // Шапка K1:N1
    const hdr = ['Настройка','Включение','Опция 1','Опция 2'];
    const hdrRange = sh.getRange(1, 11, 1, 4);
    const curHdr = hdrRange.getDisplayValues()[0].map(s => String(s).trim());
    const needsHdr = curHdr.length !== 4 || hdr.some((h,i)=> (curHdr[i]||'') !== h);
    if (needsHdr) hdrRange.setValues([hdr]);

    // Найти/создать строку
    const last = Math.max(2, sh.getLastRow());
    let row = 0;
    if (last >= 2) {
      const rows = sh.getRange(2, 11, last - 1, 4).getValues();
      for (let i=0;i<rows.length;i++) {
        const k = String(rows[i][0]||'').trim().toLowerCase();
        if (k === String(settingKey||'').trim().toLowerCase()) { row = 2 + i; break; }
      }
    }
    if (!row) {
      row = sh.getLastRow() + 1;
      sh.getRange(row, 11, 1, 4).setValues([[String(settingKey||'').trim().toLowerCase(),'','','']]);
    }

    // M/N → L
    const opt1 = String(sh.getRange(row, 13).getDisplayValue() || '').trim(); // M (left)
    const opt2 = String(sh.getRange(row, 14).getDisplayValue() || '').trim(); // N (right)
    const chosen = (side === 'left') ? opt1 : opt2;

    sh.getRange(row, 12).setValue(chosen || '');
    return chosen || '—';
  }

  // Найти верхнюю строку мерджа B:E с указанным заголовком
  function findHeaderTopRow_(sh, title) {
    const maxRow = Math.min(sh.getMaxRows(), 200);
    const norm = s => String(s || '').trim().toLowerCase();
    const want = norm(title);

    for (let r = 1; r <= maxRow; r++) {
      const rng = sh.getRange(r, 2, 1, 4); // B:E
      const val = norm(rng.getDisplayValue());
      if (!val) continue;

      if (rng.isPartOfMerge()) {
        const merged = rng.getMergedRanges()[0];
        const mval = norm(merged.getDisplayValue());
        if (mval === want) return merged.getRow();
      } else if (val === want) {
        return r;
      }
    }
    return null;
  }

  // Кнопка под «Настройки» (C:D, 2 строки) — idx 0..2
  function getSettingsButtonRange_(sh, settingsTopRow, idx) {
    const baseTop = settingsTopRow ? (settingsTopRow + 1) : 7;
    const top = baseTop + (idx * 2);
    const rng = sh.getRange(top, 3, 2, 2); // C:D
    return (rng.isPartOfMerge() ? rng.getMergedRanges()[0] : rng);
  }

  // Кнопка под «Быстрое заполнение» (E, 2 строки)
  function getQuickButtonRange_(sh, quickTopRow) {
    const top = quickTopRow ? (quickTopRow + 1) : 17; // fallback на старый кейс
    const rng = sh.getRange(top, 5, 2, 1); // E: 2 строки
    return (rng.isPartOfMerge() ? rng.getMergedRanges()[0] : rng);
  }

  function mergeAware_(rng) {
    return rng.isPartOfMerge() ? rng.getMergedRanges()[0] : rng;
  }

  // Покраска рамок кнопки C:D по стороне через Advanced Sheets API (с флагом-инвертором)
  function paintBordersBySide_(rng, side /* 'left'|'right' */) {
    const sh = rng.getSheet();
    const sheetId = sh.getSheetId();

    const r1 = rng.getRow() - 1;
    const c1 = rng.getColumn() - 1;
    const r2 = r1 + rng.getNumRows();
    const c2 = c1 + rng.getNumColumns();

    // Эквивалент XOR: активна_левая = (side==='left') !== INVERT_FLOATING_BORDER
    const leftActive = ((side === 'left') !== !!INVERT_FLOATING_BORDER);

    const leftColor  = leftActive ? KBR_COLOR_BLACK  : KBR_COLOR_EFEFEF;
    const rightColor = leftActive ? KBR_COLOR_EFEFEF : KBR_COLOR_BLACK;

    const req = {
      requests: [{
        updateBorders: {
          range: { sheetId, startRowIndex:r1, endRowIndex:r2, startColumnIndex:c1, endColumnIndex:c2 },
          left:  { style:'SOLID', width:1, color:leftColor  },
          right: { style:'SOLID', width:1, color:rightColor }
        }
      }]
    };
    Sheets.Spreadsheets.batchUpdate(req, sh.getParent().getId());
  }

  // Применить стиль к подписи (мердж в B или E) — активная/пассивная
  // Цвет и выравнивание НЕ меняем.
  function applySideStyle_(rng, isActive) {
    if (!rng) return;
    if (isActive) {
      rng.setFontWeight('bold')
         .setFontSize(FONT_ACTIVE_SIZE)
         .setBackground(BG_ACTIVE);
    } else {
      rng.setFontWeight('normal')
         .setFontSize(FONT_INACTIVE_SIZE)
         .setBackground(BG_INACTIVE);
    }
  }

  // Перестиль всех трёх опций исходя из L=M/N (рамки + подписи), без перекраски/выравнивания
  function restyleAllOptions_(sh, settingsTopRow) {
    for (let idx = 0; idx < 3; idx++) {
      const btn = getSettingsButtonRange_(sh, settingsTopRow, idx);
      if (!btn) continue;

      const side = getSideFromParams_(KBR_BTN_KEYS[idx]);

      // Рамки по стороне
      paintBordersBySide_(btn, side);

      // Подписи B/E по стороне
      const rowTop = btn.getRow();
      const rngB = mergeAware_(sh.getRange(rowTop, 2, 2, 1));
      const rngE = mergeAware_(sh.getRange(rowTop, 5, 2, 1));

      if (side === 'left') {
        applySideStyle_(rngB, true);
        applySideStyle_(rngE, false);
      } else {
        applySideStyle_(rngB, false);
        applySideStyle_(rngE, true);
      }
    }
  }

  // Переcтиль UI кнопок без расчётов (рамки и подписи)
  function restyleNow_() {
    const ss = SpreadsheetApp.getActive();
    const shCalc = ss.getSheetByName(KBR_SHEET_CALC_NAME);
    if (!shCalc) return;

    const settingsTopRow = findHeaderTopRow_(shCalc, 'Настройки');
    if (settingsTopRow) {
      restyleAllOptions_(shCalc, settingsTopRow);
    }
    // «Быстрое заполнение»: ничего не трогаем.
  }



  // Экспорт
  return {
    flip: flip,
    flipEnterPrice: flipEnterPrice,
    restyleNow: restyleNow_
  };
})();

/* ======================  ПУБЛИЧНЫЕ ФУНКЦИИ ДЛЯ КНОПОК  ====================== */
function toggleArrow_method()     { KBR_ARROWS.flip(0); }
function toggleArrow_raschet()    { KBR_ARROWS.flip(1); }
function toggleArrow_procent()    { KBR_ARROWS.flip(2); }
function toggleArrow_enterprice() { KBR_ARROWS.flipEnterPrice(); }
