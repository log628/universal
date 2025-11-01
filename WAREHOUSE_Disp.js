/** ======================================================================
 * WAREHOUSE_Disp.gs — префлайт + диалог актуализации для «🏘️ Собств. склады»
 *  - Проверяем свежесть ТОЛЬКО «Артикулы OZ»
 *  - Всегда перед расчётом запускаем Import_Sklad_GHOnly() (обновление "Доступно")
 *  - Окно показываем только если «Артикулы OZ» несвежие → обновляем getREFRESH_OZ
 *  - buildOwnWarehouses() запускаем только когда:
 *      (A) Import_Sklad_GHOnly завершён, и
 *      (B) «Артикулы OZ» свежие (изначально или после обновления)
 * ====================================================================== */

/** Точка входа: вызов из меню/кнопки */
function runWarehouseWithPreflight() {
  var cfg = WD_cfg_();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(cfg.SHEET_PARAMS);
  if (!sh) {
    SpreadsheetApp.getUi().alert('Лист «' + cfg.SHEET_PARAMS + '» не найден');
    return;
  }

  try {
    // 1) Всегда — быстрый импорт склада (только G + AH:AK)
    ss.toast('Обновление «Доступно» …', 'Склад + СС', 3);
    if (typeof Import_Sklad_GHOnly === 'function') {
      Import_Sklad_GHOnly(); // синхронно; к моменту продолжения — завершён
    } else {
      SpreadsheetApp.getUi().alert('Функция Import_Sklad_GHOnly() не найдена');
      return;
    }

    // 2) Тихий префлайт: проверяем ТОЛЬКО «Артикулы OZ»
    var now = new Date();
    var artsFresh = WD_isFresh_(sh, cfg, 'Артикулы OZ', now);

    // 3) Если свежие — сразу строим
    if (artsFresh) {
      if (typeof buildOwnWarehouses === 'function') {
        buildOwnWarehouses();
      } else {
        SpreadsheetApp.getUi().alert('Функция buildOwnWarehouses() не найдена');
      }
      return;
    }

    // 4) Если несвежие — открываем окно (в нём обновим только getREFRESH_OZ и после — построим)
    showWarehouseDispatcher();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка префлайта: ' + (e && e.message ? e.message : e));
  }
}




function runWarehouseFast() {
  var ss = SpreadsheetApp.getActive();
  try {
    // 1) Обновляем склад (как в префлайте перед окном)
    ss.toast('Обновление «Доступно» …', 'Склад + СС', 3);
    if (typeof Import_Sklad_GHOnly === 'function') {
      Import_Sklad_GHOnly(); // синхронно
    } else {
      SpreadsheetApp.getUi().alert('Функция Import_Sklad_GHOnly() не найдена');
      return;
    }

    // 2) Формируем «🏘️ Собств. склады» (без проверок свежести и без диалога)
    ss.toast('Формирование «🏘️ Собств. склады»…', 'Подготовка', 3);
    if (typeof buildOwnWarehouses === 'function') {
      buildOwnWarehouses();
      ss.toast('Готово ✅', 'Подготовка', 3);
    } else {
      SpreadsheetApp.getUi().alert('Функция buildOwnWarehouses() не найдена');
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Ошибка: ' + (e && e.message ? e.message : e));
  }
}







/** Вспомогательный: проверка свежести одного лейбла в «⚙️ Параметры» */
function WD_isFresh_(sheet, cfg, label, nowDate) {
  var row = WD_findRowByLabel_(sheet, label, cfg.RANGE_LABELS_COL);
  if (row <= 0) return false;
  var v = sheet.getRange(row, cfg.RANGE_TIMES_COL).getValue();
  var stamp = WD_normDate_(v);
  if (!stamp) return false;
  var ageHours = WD_diffHours_(stamp, nowDate);
  return ageHours <= cfg.staleHours;
}

/** Показываем диалог (только для «Артикулы OZ») */
function showWarehouseDispatcher() {
  var html = HtmlService.createTemplateFromFile('WAREHOUSE_UI')
    .evaluate()
    .setTitle('Проверка актуальности данных — Собств. склады')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Проверка актуальности данных');
}

/** Конфиг: источники и колонки — тут только «Артикулы OZ» */
function WD_cfg_() {
  return {
    SHEET_PARAMS: sheetName_('PARAMS', '⚙️ Параметры'),
    RANGE_LABELS_COL: 19 + 0, // S
    RANGE_TIMES_COL : 19 + 1, // T
    sources: [
      { key: 'arts', label: 'Артикулы OZ', runner: 'getREFRESH_OZ', expectSec: 100 }
    ],
    staleHours: 12
  };
}

function WD_normDate_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}
function WD_diffHours_(fromDate, toDate) { return (toDate.getTime() - fromDate.getTime()) / 3600000; }
function WD_fmtStamp_(d) {
  if (!d) return '';
  var dd=('0'+d.getDate()).slice(-2), mm=('0'+(d.getMonth()+1)).slice(-2), hh=('0'+d.getHours()).slice(-2), mi=('0'+d.getMinutes()).slice(-2);
  return dd+'.'+mm+' '+hh+':'+mi;
}
function WD_findRowByLabel_(sheet, label, colS) {
  var last = sheet.getLastRow(); if (last < 1) return -1;
  var rng  = sheet.getRange(1, colS, last, 1).getDisplayValues();
  var want = String(label||'').trim().toLowerCase();
  for (var r=1;r<=last;r++){ var v=String(rng[r-1][0]||'').trim().toLowerCase(); if (v===want) return r; }
  return -1;
}

/** Диалог: получить статус (только один источник) */
function WD_getStatus() {
  var cfg = WD_cfg_(), ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(cfg.SHEET_PARAMS);
  if (!sh) throw new Error('Лист «' + cfg.SHEET_PARAMS + '» не найден');
  var now = new Date(), out = [];
  cfg.sources.forEach(function (src) {
    var row = WD_findRowByLabel_(sh, src.label, cfg.RANGE_LABELS_COL);
    var stamp=null, ageHours=null, stale=true;
    if (row > 0) {
      var v = sh.getRange(row, cfg.RANGE_TIMES_COL).getValue();
      stamp = WD_normDate_(v);
      if (stamp) { ageHours = WD_diffHours_(stamp, now); stale = ageHours > cfg.staleHours; }
    }
    out.push({
      key: src.key, label: src.label, runner: src.runner, expectSec: src.expectSec, row: row,
      stampIso: stamp ? stamp.toISOString() : null,
      stampHuman: stamp ? WD_fmtStamp_(stamp) : '',
      ageHours: (ageHours != null) ? Math.floor(ageHours) : null,
      isStale: !!stale
    });
  });
  return { nowIso: now.toISOString(), sources: out, staleHours: cfg.staleHours };
}

/** Диалог: запустить обновление «Артикулы OZ» и поставить штамп T */
function WD_runSourceUpdate(key) {
  var cfg = WD_cfg_(), ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(cfg.SHEET_PARAMS);
  if (!sh) throw new Error('Лист «' + cfg.SHEET_PARAMS + '» не найден');
  var src = cfg.sources.find(function (s){ return s.key===key; });
  if (!src) throw new Error('Неизвестный источник: ' + key);
  var row = WD_findRowByLabel_(sh, src.label, cfg.RANGE_LABELS_COL);
  if (row <= 0) throw new Error('Строка «' + src.label + '» в колонке S не найдена');

  var fn = this[src.runner];
  if (typeof fn !== 'function') throw new Error('Функция "'+src.runner+'" не найдена');

  // Обновление «Артикулы OZ»
  fn();

  // Штамп свежести в T
  var now = new Date();
  sh.getRange(row, cfg.RANGE_TIMES_COL).setValue(now);
  return { key: key, stampIso: now.toISOString(), stampHuman: WD_fmtStamp_(now), ageHours: 0 };
}

/** Диалог: финальный запуск формирования склада */
function WD_runBuildWarehouse() {
  if (typeof buildOwnWarehouses !== 'function') throw new Error('Функция buildOwnWarehouses не найдена');
  buildOwnWarehouses();
  return { ok:true, doneAtIso: new Date().toISOString() };
}

// include() — если нужно подключать partial'ы в HTML
function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Если sheetName_ не объявлен где-то ещё — оставим совместимый помощник */
function sheetName_(key, fallback) {
  try {
    if (typeof REF !== 'undefined' && REF.SHEETS && REF.SHEETS[key]) return REF.SHEETS[key];
  } catch (_) {}
  return fallback;
}
