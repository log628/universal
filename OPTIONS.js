/* ======================  KBR_ARROWS (изолированный модуль)  ====================== */
/* Публичные функции для привязки к рисункам:
   toggleArrow_method / toggleArrow_raschet / toggleArrow_procent / toggleArrow_enterprice

   ДИНАМИЧЕСКИЕ ПОЗИЦИИ:
   - «Настройки» = мердж B:E с текстом "Настройки". Под ним 3 кнопки, каждая 2 строки, в C:D:
       idx 0 → Метод, idx 1 → Расчёт, idx 2 → Процент
   - «Быстрое заполнение» = мердж B:E с текстом "Быстрое заполнение".
       Под ним кнопка в колонке E, 2 строки (E{row+1}:E{row+2}).
*/

var KBR_ARROWS = KBR_ARROWS || (function () {
  // ===== Константы листов
  const KBR_SHEET_CALC_NAME   = (typeof SHEET_CALC   === 'string' ? SHEET_CALC   : '⚖️ Калькулятор');
  const KBR_SHEET_PARAMS_NAME = (typeof SHEET_PARAMS === 'string' ? SHEET_PARAMS : '⚙️ Параметры');

  // ===== Стрелки/цвета
  const KBR_ARROW_L = '◀️';
  const KBR_ARROW_R = '▶️';

  const KBR_COLOR_BLACK  = { red: 0, green: 0, blue: 0 };
  const KBR_COLOR_EFEFEF = { red: 239/255, green: 239/255, blue: 239/255 };

  // ===== Поддерживаемые «настройки» (3 стрелки под «Настройки»)
  const KBR_BTN_KEYS     = ['метод','расчет','процент'];
  const KBR_BTN_LABELSUI = ['Метод','Расчёт','Процент'];

  // ===== Ключ для «ввод цены» (под «Быстрое заполнение»)
  const KBR_ENTERPRICE_LABEL_KEY = 'ввод цены';
  const KBR_ENTERPRICE_LABEL_UI  = 'Ввод цены';

  // --------------------------------------------------------------------------------
  //                                  ПУБЛИЧНЫЕ
  // --------------------------------------------------------------------------------

  // Универсальный переключатель «Настройки»: idx 0..2
  function flip(idx) {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(KBR_SHEET_CALC_NAME);
    if (!sh) return;

    // 1) найдём верх «Настройки»
    const settingsRow = findHeaderTopRow_(sh, 'Настройки');
    const rng = getSettingsButtonRange_(sh, settingsRow, idx); // C:D, 2 строки
    if (!rng) {
      ss.toast('⚠️ Не найден блок «Настройки»', 'Ошибка', 4);
      return;
    }

    // 2) Переключаем стрелку
    const cur = String(rng.getCell(1,1).getDisplayValue() || '').trim();
    const nextArrow =
      (cur === KBR_ARROW_R) ? KBR_ARROW_L :
      (cur === KBR_ARROW_L) ? KBR_ARROW_R :
      KBR_ARROW_R; // дефолт вправо
    rng.setValue(nextArrow);
    paintSidesExact_(rng, nextArrow);

    // 3) Запись выбранной опции в «⚙️ Параметры» и тост
    const chosenValue = writeParams_(KBR_BTN_KEYS[idx], nextArrow);
    ss.toast('📃 ' + KBR_BTN_LABELSUI[idx] + ' = ' + (chosenValue || '—'), 'Готово', 3);
  }

  // «Ввод цены» (под «Быстрое заполнение»):
  // если L("ввод цены") = "Заполнено" — чистим H4:H{по G}; иначе — заполняем из «⛓️ Параллель».
  // Источник подстановки: если L("процент") = "Фактич. %" → Q, иначе → B
  // Доп. условие: заполняем ТОЛЬКО если в «⛓️ Параллель» M != "нет СС"
  function flipEnterPrice() {
    const ss = SpreadsheetApp.getActive();
    const shCalc = ss.getSheetByName(KBR_SHEET_CALC_NAME);
    const shPar  = ss.getSheetByName(KBR_SHEET_PARAMS_NAME);
    if (!shCalc || !shPar) return;

    // 1) читаем «ввод цены» и «процент» из «⚙️ Параметры» (K..N)
    const lastPar = Math.max(2, shPar.getLastRow());
    const rowsPar = (lastPar >= 2) ? shPar.getRange(2, 11, lastPar - 1, 4).getDisplayValues() : []; // K..N
    let stateEnterPrice = ''; // L по "ввод цены"
    let statePercent    = ''; // L по "процент"
    for (let i = 0; i < rowsPar.length; i++) {
      const key = String(rowsPar[i][0] || '').trim().toLowerCase(); // K
      if (key === 'ввод цены') stateEnterPrice = String(rowsPar[i][1] || '').trim(); // L
      if (key === 'процент')    statePercent    = String(rowsPar[i][1] || '').trim(); // L
    }

    // 2) рабочая высота по последней непустой в G4:G
    const lastRowCalc = shCalc.getLastRow();
    const height      = Math.max(lastRowCalc - 3, 1); // 4..lastRow
    const gVals       = shCalc.getRange(4, 7, height, 1).getValues(); // G4:G
    let lastIdx = -1;
    for (let i = gVals.length - 1; i >= 0; i--) {
      if (String(gVals[i][0]).trim() !== '') { lastIdx = i; break; }
    }
    if (lastIdx < 0) {
      ss.toast('📝 Ввод цены: заполнено', 'Готово', 3);
      return;
    }
    const endRow = 4 + lastIdx;          // включительно
    const num    = endRow - 3;           // кол-во строк
    const hRange = shCalc.getRange(4, 8, num, 1); // H4:H{endRow}

    if (stateEnterPrice === 'Заполнено') {
      // === ОЧИСТКА ===
      hRange.clearContent();
      hRange.setNumberFormat('#,##0');
      hRange.setHorizontalAlignment('center');
      hRange.setVerticalAlignment('middle');
      ss.toast('🧹 Ввод цены: очищено', 'Готово', 3);
    } else {
      // === ЗАПОЛНЕНИЕ ===
      const shParallel = ss.getSheetByName('⛓️ Параллель');
      if (!shParallel) {
        ss.toast('📝 Ввод цены: заполнено', 'Готово', 3);
        return;
      }
      const lastParRow = shParallel.getLastRow();
      if (lastParRow < 2) {
        ss.toast('📝 Ввод цены: заполнено', 'Готово', 3);
        return;
      }

      const srcColIndex = (statePercent === 'Фактич. %') ? 17 : 2; // Q=17, B=2

      // Забираем A (артикулы), M (флаг СС), и srcColIndex (значения)
      const parA = shParallel.getRange(2, 1,  lastParRow - 1, 1).getValues();           // A2:A
      const parM = shParallel.getRange(2, 13, lastParRow - 1, 1).getDisplayValues();    // M2:M
      const parV = shParallel.getRange(2, srcColIndex, lastParRow - 1, 1).getValues();  // B2:B или Q2:Q

      // Строим карту только для тех, где M != "нет СС"
      const map = new Map(); // артикул -> значение
      for (let i = 0; i < parA.length; i++) {
        const art = String(parA[i][0] || '').trim();
        if (!art) continue;
        const mark = String(parM[i][0] || '').trim().toLowerCase();
        if (mark === 'нет сс') continue; // пропускаем — НЕ заполняем
        map.set(art, parV[i][0]);
      }

      const hCurr = hRange.getValues();
      const out = new Array(num);
      for (let i = 0; i < num; i++) {
        const art = String(gVals[i][0] || '').trim();
        if (map.has(art)) {
          let v = map.get(art);
          const n = (typeof v === 'number') ? v : Number(String(v).replace(',', '.'));
          out[i] = [Number.isFinite(n) ? n : ''];
        } else {
          out[i] = [hCurr[i][0]]; // оставить как было
        }
      }

      hRange.setValues(out);
      hRange.setNumberFormat('#,##0');
      hRange.setHorizontalAlignment('center');
      hRange.setVerticalAlignment('middle');

      ss.toast('📝 Ввод цены: заполнено', 'Готово', 3);
    }

    // НИКАКИХ проверок busy/cooldown и изменений внешнего вида «кнопки» не делаем.
  }

  // --------------------------------------------------------------------------------
  //                               ВСПОМОГАТЕЛЬНЫЕ
  // --------------------------------------------------------------------------------

  // Находим верхнюю строку мерджа B:E с нужным текстом (без учёта регистра/пробелов по краям)
  function findHeaderTopRow_(sh, title) {
    const maxRow = Math.min(sh.getMaxRows(), 200); // ограничимся верхней частью
    const norm = s => String(s || '').trim().toLowerCase();
    const want = norm(title);

    for (let r = 1; r <= maxRow; r++) {
      const rng = sh.getRange(r, 2, 1, 4); // B:E
      const val = norm(rng.getDisplayValue());
      if (!val) continue;
      // проверяем, что это именно мердж и текст совпал
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

  // Возвращает диапазон-мердж кнопки под «Настройки» для idx 0..2 (C:D, 2 строки)
  function getSettingsButtonRange_(sh, settingsTopRow, idx) {
    // fallback к прежним координатам, если заголовок не нашли
    const baseTop = settingsTopRow ? (settingsTopRow + 1) : 7;
    const top = baseTop + (idx * 2);
    const rng = sh.getRange(top, 3, 2, 2); // C:D (2 строки)
    // если там уже мердж — вернём его, иначе сам диапазон
    return (rng.isPartOfMerge() ? rng.getMergedRanges()[0] : rng);
  }

  // Возвращает диапазон-мердж кнопки под «Быстрое заполнение» (E, 2 строки)
  function getQuickButtonRange_(sh, quickTopRow) {
    const top = quickTopRow ? (quickTopRow + 1) : 17; // fallback на старый кейс
    const rng = sh.getRange(top, 5, 2, 1); // E: 2 строки
    return (rng.isPartOfMerge() ? rng.getMergedRanges()[0] : rng);
  }

  // Точная прокраска краёв (лев/прав) для ◀️/▶️ через Advanced Sheets API
  function paintSidesExact_(rng, arrow) {
    const sh = rng.getSheet();
    const sheetId = sh.getSheetId();

    const r1 = rng.getRow() - 1;
    const c1 = rng.getColumn() - 1;
    const r2 = r1 + rng.getNumRows();
    const c2 = c1 + rng.getNumColumns();

    const leftColor  = (arrow === KBR_ARROW_R) ? KBR_COLOR_BLACK  : KBR_COLOR_EFEFEF;
    const rightColor = (arrow === KBR_ARROW_R) ? KBR_COLOR_EFEFEF : KBR_COLOR_BLACK;

    const req = {
      requests: [
        {
          updateBorders: {
            range: {
              sheetId: sheetId,
              startRowIndex: r1,
              endRowIndex:   r2,
              startColumnIndex: c1,
              endColumnIndex:   c2
            },
            left:  { style: 'SOLID', width: 1, color: leftColor  },
            right: { style: 'SOLID', width: 1, color: rightColor }
          }
        }
      ]
    };
    Sheets.Spreadsheets.batchUpdate(req, sh.getParent().getId());
  }

  // Запись выбранной опции в «⚙️ Параметры» (три строки: метод/расчет/процент)
  function writeParams_(settingName, arrow) {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(KBR_SHEET_PARAMS_NAME) || ss.insertSheet(KBR_SHEET_PARAMS_NAME);

    // K1:N1 — шапка
    const hdr = ['Настройка', 'Включение', 'Опция 1', 'Опция 2'];
    const hdrRange = sh.getRange(1, 11, 1, 4);
    const curHdr = hdrRange.getDisplayValues()[0].map(s => String(s).trim());
    const needsHdr = curHdr.length !== 4 || hdr.some((h, i) => (curHdr[i] || '') !== h);
    if (needsHdr) hdrRange.setValues([hdr]);

    const controllable = KBR_BTN_KEYS; // ['метод','расчет','процент']
    const need = {}; for (let i = 0; i < controllable.length; i++) need[controllable[i]] = true;

    const last = Math.max(2, sh.getLastRow());
    const data = (last >= 2) ? sh.getRange(2, 11, last - 1, 4).getValues() : [];
    const rowByName = {};

    for (let r = 0; r < data.length; r++) {
      const kVal = String(data[r][0] || '').trim().toLowerCase();
      if (need[kVal]) rowByName[kVal] = 2 + r;
    }

    const toAppend = [];
    for (let j = 0; j < controllable.length; j++) {
      const nm = controllable[j];
      if (!rowByName[nm]) toAppend.push([nm, '', '', '']);
    }
    if (toAppend.length) {
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 11, toAppend.length, 4).setValues(toAppend);
      for (let a = 0; a < toAppend.length; a++) {
        rowByName[toAppend[a][0]] = start + a;
      }
    }

    const key = String(settingName).trim().toLowerCase();
    if (!need[key]) return '—';

    const row = rowByName[key];
    if (!row) return '—';

    const valOpt1 = sh.getRange(row, 13).getDisplayValue(); // M — Опция 1
    const valOpt2 = sh.getRange(row, 14).getDisplayValue(); // N — Опция 2
    const chosen  = (arrow === KBR_ARROW_L) ? valOpt1 : valOpt2;

    sh.getRange(row, 12).setValue(chosen); // L — «Включение»
    return chosen || '—';
  }

  // Экспорт
  return {
    flip: flip,
    flipEnterPrice: flipEnterPrice
  };
})();

/* ======================  ПУБЛИЧНЫЕ ФУНКЦИИ ДЛЯ КНОПОК  ====================== */
function toggleArrow_method()     { KBR_ARROWS.flip(0); }
function toggleArrow_raschet()    { KBR_ARROWS.flip(1); }
function toggleArrow_procent()    { KBR_ARROWS.flip(2); }
function toggleArrow_enterprice() { KBR_ARROWS.flipEnterPrice(); }
