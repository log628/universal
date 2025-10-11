function showForecastDispatcher() {
  var html = HtmlService.createTemplateFromFile('FORCAST_UI') // ← новое имя HTML-файла
    .evaluate()
    .setTitle('Проверка актуальности данных — Форкаст')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Проверка актуальности данных');
}

/** Конфиг источников и прочее */
function FD_cfg_() {
  return {
    SHEET_PARAMS: sheetName_('PARAMS', '⚙️: Параметры'),
    RANGE_LABELS_COL: 19 + 0, // S
    RANGE_TIMES_COL : 19 + 1, // T
    sources: [
      { key: 'ozon',  label: 'Физ.оборот OZ', runner: 'getFiz_OZ',    expectSec: 20 }, // было 15
      { key: 'wb',    label: 'Физ.оборот WB', runner: 'getFiz_WB',    expectSec: 12 }, // было 10
      { key: 'sklad', label: 'Склад + СС',    runner: 'Import_Sklad', expectSec: 8  }, // было 5
    ],
    staleHours: 12
  };
}

function sheetName_(key, fallback) {
  try {
    if (typeof REF !== 'undefined' && REF.SHEETS && REF.SHEETS[key]) return REF.SHEETS[key];
  } catch (_) {}
  return fallback;
}

function FD_normDate_(v) { if (!v) return null; if (v instanceof Date && !isNaN(v.getTime())) return v; var d=new Date(v); return isNaN(d.getTime())?null:d; }
function FD_diffHours_(fromDate, toDate) { return (toDate.getTime() - fromDate.getTime()) / 3600000; }
function FD_fmtStamp_(d) { if (!d) return ''; var dd=('0'+d.getDate()).slice(-2), mm=('0'+(d.getMonth()+1)).slice(-2), hh=('0'+d.getHours()).slice(-2), mi=('0'+d.getMinutes()).slice(-2); return dd+'.'+mm+' '+hh+':'+mi; }
function FD_findRowByLabel_(sheet, label, colS) {
  var last = sheet.getLastRow(); if (last < 1) return -1;
  var rng  = sheet.getRange(1, colS, last, 1).getDisplayValues();
  var want = String(label||'').trim().toLowerCase();
  for (var r=1;r<=last;r++){ var v=String(rng[r-1][0]||'').trim().toLowerCase(); if (v===want) return r; }
  return -1;
}

function FD_getStatus() {
  var cfg = FD_cfg_(), ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(cfg.SHEET_PARAMS);
  if (!sh) throw new Error('Лист «' + cfg.SHEET_PARAMS + '» не найден');
  var now = new Date(), out = [];
  cfg.sources.forEach(function (src) {
    var row = FD_findRowByLabel_(sh, src.label, cfg.RANGE_LABELS_COL);
    var stamp=null, ageHours=null, stale=true;
    if (row > 0) {
      var v = sh.getRange(row, cfg.RANGE_TIMES_COL).getValue();
      stamp = FD_normDate_(v);
      if (stamp) { ageHours = FD_diffHours_(stamp, now); stale = ageHours > cfg.staleHours; }
    }
    out.push({
      key: src.key, label: src.label, runner: src.runner, expectSec: src.expectSec, row: row,
      stampIso: stamp ? stamp.toISOString() : null,
      stampHuman: stamp ? FD_fmtStamp_(stamp) : '',
      ageHours: (ageHours != null) ? Math.floor(ageHours) : null,
      isStale: !!stale
    });
  });
  return { nowIso: now.toISOString(), sources: out, staleHours: cfg.staleHours };
}

function FD_runSourceUpdate(key) {
  var cfg = FD_cfg_(), ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(cfg.SHEET_PARAMS);
  if (!sh) throw new Error('Лист «' + cfg.SHEET_PARAMS + '» не найден');
  var src = cfg.sources.find(function (s){ return s.key===key; });
  if (!src) throw new Error('Неизвестный источник: ' + key);
  var row = FD_findRowByLabel_(sh, src.label, cfg.RANGE_LABELS_COL);
  if (row <= 0) throw new Error('Строка «' + src.label + '» в колонке S не найдена');
  var fn = this[src.runner];
  if (typeof fn !== 'function') throw new Error('Функция "'+src.runner+'" не найдена');
  fn(); // реальное обновление
  var now = new Date();
  sh.getRange(row, cfg.RANGE_TIMES_COL).setValue(now);
  return { key, stampIso: now.toISOString(), stampHuman: FD_fmtStamp_(now), ageHours: 0 };
}

function FD_runBuildForecast() {
  if (typeof buildForecast_All !== 'function') throw new Error('Функция buildForecast_All не найдена');
  buildForecast_All();
  return { ok:true, doneAtIso: new Date().toISOString() };
}

// хелпер для include(), если решишь подключать отдельные partial'ы
function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
