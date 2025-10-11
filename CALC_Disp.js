/** ======================================================================
 * CALC_Disp.gs — диспетчер актуализации источников перед КАЛЬКУЛЯТОРОМ
 * Логика:
 *   - Тихий префлайт (только «Цены OZ», «Цены WB») → если ОБА свежие, даём
 *     продолжить runLayoutImmediate() без окна; иначе — показываем окно.
 *   - В окне проверяем уже три источника (добавляется «Склад + СС»),
 *     обновляем устаревшие и затем запускаем рассчёт.
 * ====================================================================== */

/** Открыть окно-диспетчер */
function showCalcDispatcher() {
  var html = HtmlService.createTemplateFromFile('CALC_UI')
    .evaluate()
    .setTitle('Проверка актуальности данных — Калькулятор')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Проверка актуальности данных');
}

/** Конфиг источников и общие параметры */
function CD_cfg_() {
  return {
    SHEET_PARAMS: sheetName_('PARAMS', '⚙️ Параметры'),
    RANGE_LABELS_COL: 19, // S
    RANGE_TIMES_COL : 20, // T (штамп)
    sources: [
      { key: 'price_oz', label: 'Цены OZ',  runner: 'getREFRESHprices_OZ', expectSec: 40 },
      { key: 'price_wb', label: 'Цены WB',  runner: 'getREFRESHprices_WB', expectSec: 15 },
      { key: 'sklad',    label: 'Склад + СС', runner: 'Import_Sklad',       expectSec: 10 },
    ],
    staleHours: 12
  };
}

/** Имя листа из REF либо фолбэк */
function sheetName_(key, fallback) {
  try {
    if (typeof REF !== 'undefined' && REF.SHEETS && REF.SHEETS[key]) return REF.SHEETS[key];
  } catch (_) {}
  return fallback;
}

/** Утилиты времени/поиска строк */
function CD_normDate_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  // ожидаемый формат строки из REF.logRun(): "dd.MM HH:mm" (без года)
  var s = String(v||'').trim();
  var m = /^(\d{2})\.(\d{2})\s+(\d{2}):(\d{2})$/.exec(s);
  if (m) {
    var ss = SpreadsheetApp.getActive();
    var tz = ss.getSpreadsheetTimeZone() || 'Etc/GMT';
    var year = Number(Utilities.formatDate(new Date(), tz, 'yyyy'));
    var d = new Date(year, Number(m[2])-1, Number(m[1]), Number(m[3]), Number(m[4]), 0, 0);
    return isNaN(d.getTime()) ? null : d;
  }
  var d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}
function CD_diffHours_(fromDate, toDate) { return (toDate.getTime() - fromDate.getTime()) / 3600000; }
function CD_fmtStamp_(d) {
  if (!d) return '';
  var dd=('0'+d.getDate()).slice(-2), mm=('0'+(d.getMonth()+1)).slice(-2);
  var hh=('0'+d.getHours()).slice(-2), mi=('0'+d.getMinutes()).slice(-2);
  return dd+'.'+mm+' '+hh+':'+mi;
}
function CD_findRowByLabel_(sheet, label, colS) {
  var last = sheet.getLastRow(); if (last < 1) return -1;
  var rng  = sheet.getRange(1, colS, last, 1).getDisplayValues();
  var want = String(label||'').trim().toLowerCase();
  for (var r=1;r<=last;r++){
    var v=String(rng[r-1][0]||'').trim().toLowerCase();
    if (v===want) return r;
  }
  return -1;
}

/** Внутренняя проверка свежести по метке (S→T) */
function CD_isFreshByLabel_(label, staleHours) {
  var cfg = CD_cfg_();
  var ss  = SpreadsheetApp.getActive();
  var sh  = ss.getSheetByName(cfg.SHEET_PARAMS);
  if (!sh) return false;
  var row = CD_findRowByLabel_(sh, label, cfg.RANGE_LABELS_COL);
  if (row <= 0) return false;
  var v   = sh.getRange(row, cfg.RANGE_TIMES_COL).getValue();
  var dt  = CD_normDate_(v);
  if (!dt) return false;
  var age = CD_diffHours_(dt, new Date());
  return age <= (staleHours || cfg.staleHours);
}

/** Публичный префлайт: вернуть true — можно считать без окна; false — окно будет показано и расчёт не продолжаем. */
function CALC_preflightOrShowDialog() {
  var cfg = CD_cfg_();
  var bothFresh =
    CD_isFreshByLabel_('Цены OZ', cfg.staleHours) &&
    CD_isFreshByLabel_('Цены WB', cfg.staleHours);
  if (bothFresh) return true; // безопасно продолжать runLayoutImmediate()

  // хотя бы одна «цена» просрочена — показываем диспетчер
  showCalcDispatcher();
  return false;
}

/** Статусы для UI: три источника */
function CD_getStatus() {
  var cfg = CD_cfg_(), ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(cfg.SHEET_PARAMS);
  if (!sh) throw new Error('Лист «' + cfg.SHEET_PARAMS + '» не найден');

  var now = new Date(), out = [];
  cfg.sources.forEach(function (src) {
    var row = CD_findRowByLabel_(sh, src.label, cfg.RANGE_LABELS_COL);
    var stamp=null, ageHours=null, stale=true;
    if (row > 0) {
      var v = sh.getRange(row, cfg.RANGE_TIMES_COL).getValue();
      stamp = CD_normDate_(v);
      if (stamp) { ageHours = CD_diffHours_(stamp, now); stale = ageHours > cfg.staleHours; }
    }
    out.push({
      key: src.key, label: src.label, runner: src.runner, expectSec: src.expectSec, row: row,
      stampIso: stamp ? stamp.toISOString() : null,
      stampHuman: stamp ? CD_fmtStamp_(stamp) : '',
      ageHours: (ageHours != null) ? Math.floor(ageHours) : null,
      isStale: !!stale
    });
  });
  return { nowIso: now.toISOString(), sources: out, staleHours: cfg.staleHours };
}

/** Обновить конкретный источник (runner уже сам пишет штамп через REF.logRun) */
function CD_runSourceUpdate(key) {
  var cfg = CD_cfg_(), ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(cfg.SHEET_PARAMS);
  if (!sh) throw new Error('Лист «' + cfg.SHEET_PARAMS + '» не найден');
  var src = cfg.sources.find(function (s){ return s.key===key; });
  if (!src) throw new Error('Неизвестный источник: ' + key);

  var row = CD_findRowByLabel_(sh, src.label, cfg.RANGE_LABELS_COL);
  if (row <= 0) throw new Error('Строка «' + src.label + '» в колонке S не найдена');

  var fn = this[src.runner];
  if (typeof fn !== 'function') throw new Error('Функция "'+src.runner+'" не найдена');

  // обновление
  fn();

  // читаем актуальный штамп из T (runner уже записал)
  var raw = sh.getRange(row, cfg.RANGE_TIMES_COL).getValue();
  var d   = CD_normDate_(raw);
  var now = new Date();
  return {
    key: key,
    stampIso: d ? d.toISOString() : null,
    stampHuman: d ? CD_fmtStamp_(d) : '',
    ageHours: d ? Math.max(0, Math.floor(CD_diffHours_(d, now))) : null
  };
}

/** Финальный запуск калькулятора после актуализации (используем текущий кабинет из REF) */
function CD_runBuildCalculator() {
  if (typeof runLayoutImmediate !== 'function') throw new Error('Функция runLayoutImmediate не найдена');

  var cab = '';
  try { cab = (REF && REF.getCabinetControlValue) ? REF.getCabinetControlValue() : ''; } catch(_){}
  cab = String(cab||'').trim();

  if (!cab || cab === '<выберите кабинет>') {
    throw new Error('Не выбран кабинет (именованный диапазон muff_cabs)');
  }

  // Запускаем как обычно: префлайт внутри runLayoutImmediate второй раз окно не поднимет (цены уже свежие).
  runLayoutImmediate(cab);
  return { ok:true, doneAtIso: new Date().toISOString() };
}

/** include_ для html partial'ов (если потребуется) */
function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
