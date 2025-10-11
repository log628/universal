function getREFRESH_OZ() {
  var ss = SpreadsheetApp.getActive();
  var T0 = Date.now();
  var log = function (label, extra) {
    var ms = String(Date.now() - T0).padStart(7, ' ');
    console.log('[' + ms + ' ms] getREFRESH_OZ | ' + label + (extra ? ' | ' + extra : ''));
  };

  log('START');

  var DST_SHEET   = REF.SHEETS.ARTS_OZ;
  var PARAM_SHEET = REF.SHEETS.PARAMS;

  var TOTAL_COLS = 13; // A..M
  var HEADERS = REF.getArtsHeaders('OZ');

  // ===== 1) Префиксы «Раздел» и «Своя категория» (P:Q)
  var sections = (function readSections_() {
    var out = [];
    var sh = ss.getSheetByName(PARAM_SHEET);
    if (!sh) return out;
    var last = sh.getLastRow();
    if (last < 2) return out;
    var vals = sh.getRange(2, 16, last - 1, 2).getDisplayValues(); // P:Q
    for (var i = 0; i < vals.length; i++) {
      var pref = String(vals[i][0] || '').trim().toLowerCase();
      var own  = String(vals[i][1] || '').trim();
      if (pref) out.push({ prefix: pref, ownCat: own });
    }
    return out;
  })();
  log('SECTIONS', 'count=' + sections.length);

  if (sections.length === 0) {
    var sh0 = ensureLayoutN_(ss, DST_SHEET, HEADERS, TOTAL_COLS);
    paintHeaderBlocks_OZ_(sh0);
    trimRowsAfter_(sh0, 1);
    ss.toast('Нет префиксов в ⚙️ Параметры!P — выгрузка пуста', 'Готово', 5);
    log('END', 'no prefixes -> 0 rows');
    try { REF.logRun('Артикулы OZ', [], 'OZ'); } catch(_){}
  triggerRebuildIfOZ_();                 // ← ДОБАВИТЬ

    return;
  }

  // ===== 2) Выбранные OZ-кабинеты (чекбокс H)
  var pick = getSelectedCabinetsFromParams_OZ_(); // { list, total }
  var selected = pick.list || [];

  var sh = ensureLayoutN_(ss, DST_SHEET, HEADERS, TOTAL_COLS);
  paintHeaderBlocks_OZ_(sh);
  clearDataUnderHeader_(sh, TOTAL_COLS);

  if (!selected.length) {
    trimRowsAfter_(sh, 1);
    ss.toast('В ⚙️ Параметры ни один OZON-кабинет не отмечен', 'Нет выбора', 5);
    log('END', 'no selected cabinets');
    try { REF.logRun('Артикулы OZ', [], 'OZ'); } catch(_){}

      triggerRebuildIfOZ_();                 // ← ДОБАВИТЬ
    return;
  }

  var colorMap = REF.readCabinetColorMap('OZON');

  // ===== 3) Сборка строк по кабинету
  function buildRowsForCabinet(cabKey) {
    ss.toast('Запрос данных кабинета ' + cabKey, 'Выполнение', 3);
    var api = new OZONAPI(cabKey);

    var reportId = api.makeArtsReport();
    Utilities.sleep(1000);
    var artsObj  = api.getArts(reportId) || {}; // { offerId: CSVcols }
    var offers   = Object.keys(artsObj);
    if (!offers.length) return [];

    var details = api.getArtsDetails(offers) || []; // items[]
    var detMap  = details.reduce(function (acc, it) {
      var offer = it.offer_id || it.offer || '';
      if (offer) acc[offer] = it;
      return acc;
    }, {});

    var missingForSku = offers.filter(function (of) {
      var d = detMap[of];
      return !(d && (d.sku || d.sku_id || d.id));
    });
    var skuFallbackMap = {};
    if (missingForSku.length) {
      try { skuFallbackMap = api.getSkusByOffers(missingForSku) || {}; } catch (_) { skuFallbackMap = {}; }
    }

var cats = {};
try { cats = api.getCategoryTreeRU() || {}; } catch(_) { cats = {}; }
var byCategoryId     = cats.byCategoryId || {};
var byTypeId         = cats.byTypeId || {};
var typeNameByTypeId = cats.typeNameByTypeId || {};

function pickCategoryAndType_(det) {
  var cid = det && (det.category_id || det.description_category_id);
  var tid = det && det.type_id;

  var category = '';
  if (cid != null && byCategoryId[String(cid)]) {
    category = byCategoryId[String(cid)];
  } else if (tid != null && byTypeId[String(tid)]) {
    category = byTypeId[String(tid)];
  }

  var typeName = (tid != null && typeNameByTypeId[String(tid)]) ? typeNameByTypeId[String(tid)] : '';
  return { category: category, typeName: typeName };
}

function pickCommissions_(det) {
  var fboPct = '', fbsPct = '', rfbsPct = '';
  var arr = det && Array.isArray(det.commissions) ? det.commissions : [];

  // универсальный парсер числа (если есть REF.toNumber — используем его)
  var toNum = (typeof REF !== 'undefined' && typeof REF.toNumber === 'function')
    ? REF.toNumber
    : function(v){ return Number(String(v).replace(',', '.')); };

  for (var i = 0; i < arr.length; i++) {
    var c = arr[i];
    var v = (c && c.percent != null) ? toNum(c.percent) : NaN;
    if (!isNaN(v)) {
      v = v + 5; // ← добавляем +5 п.п. ДЛЯ ВСЕХ схем
      if (c.sale_schema === 'FBO')  fboPct  = v;
      if (c.sale_schema === 'FBS')  fbsPct  = v;
      if (c.sale_schema === 'RFBS') rfbsPct = v;
    }
  }
  return { fboPct: fboPct, fbsPct: fbsPct, rfbsPct: rfbsPct };
}




    function pickSku_(det, offer) {
      return (det && (det.sku || det.sku_id || det.id)) || skuFallbackMap[offer] || '';
    }
    function calcVolumeLitersFromCsv_(csv) {
      return String(csv['Объем товара, л'] || '').replace(/\./g, ',');
    }
    function pickPriceFromCsv_(csv) {
      var priceStr = String(csv['Текущая цена с учетом скидки, ₽'] || '0');
      var priceNum = REF.toNumber(priceStr);
      return String(priceNum).replace('.', ',');
    }
    function resolveSection_(offer) {
      var tail = String(offer || '').toLowerCase();
      tail = tail.length >= 4 ? tail.substring(3) : '';
      if (!tail) return null;
      for (var i = 0; i < sections.length; i++) {
        var p = sections[i];
        if (tail.indexOf(p.prefix) === 0) {
          return { section: 'X', ownCat: p.ownCat || '' };
        }
      }
      return null;
    }

    var rows = [];
    for (var i = 0; i < offers.length; i++) {
      var offer = offers[i];
      var csv   = artsObj[offer] || {};
      var det   = detMap[offer] || {};
      var sec   = resolveSection_(offer);
      if (!sec) continue;

var cm = pickCommissions_(det);
var ct = pickCategoryAndType_(det);
var catCell = ct.category || '';
if (ct.typeName) catCell += ' | ' + ct.typeName;

rows.push([
  cabKey,                                        // A
  offer,                                         // B
  csv['Отзывы'] || '',                           // C
  String(csv['Рейтинг'] || '').replace(/\./g,','), // D
  catCell,                                       // E  <-- теперь "Категория | Тип"
  cm.fboPct,                                     // F
  cm.fbsPct,                                     // G
  cm.rfbsPct,                                    // H
  calcVolumeLitersFromCsv_(csv),                 // I
  pickPriceFromCsv_(csv),                        // J
  pickSku_(det, offer),                          // K
  'X',                                           // L
  sec.ownCat || ''                               // M
]);

    }
    return rows;
  }

  // ===== 4) Обход кабинетов, запись и стили
  var allRows = [];
  var successCabs = [];
  for (var i = 0; i < selected.length; i++) {
    var cab = selected[i];
    var tCab = Date.now();
    var rows = [];
    try { rows = buildRowsForCabinet(cab) || []; } catch (e) {
      log('WARN buildRowsForCabinet', cab + ': ' + ((e && e.message) || e));
    }
    if (rows.length) successCabs.push(cab);
    allRows = allRows.concat(rows);
    log('CAB done', cab + ', rowsNow=' + allRows.length + ', took=' + (Date.now() - tCab) + 'ms');
  }

  clearDataUnderHeader_(sh, TOTAL_COLS);
  if (allRows.length) {
    sh.getRange(2, 1, allRows.length, TOTAL_COLS).setValues(allRows);
    var fills = makeRowFillsFromCabinet_(allRows, colorMap, TOTAL_COLS);
    if (fills) sh.getRange(2, 1, allRows.length, TOTAL_COLS).setBackgrounds(fills);
    applyPostStyles_(sh, 2, allRows.length, TOTAL_COLS);
  }
  trimRowsAfter_(sh, (allRows.length || 0) + 1);

  ss.toast('Готово! Обновлено ' + (allRows.length || 0) + ' строк (OZON)', DST_SHEET, 5);
  log('END', 'totalRows=' + (allRows.length || 0));

  // ===== 5) ЛОГ — только успешные кабинеты
try { REF.logRun('Артикулы OZ', successCabs, 'OZON'); } catch(_){}

triggerRebuildIfOZ_();                   // ← ДОБАВИТЬ

  /* ==================== локальные хелперы ==================== */
  function ensureLayoutN_(ss, sheetName, headers, N) {
    var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    var maxCols = sh.getMaxColumns();
    if (maxCols > N) sh.deleteColumns(N + 1, maxCols - N);
    if (sh.getMaxColumns() < N) sh.insertColumnsAfter(sh.getMaxColumns(), N - sh.getMaxColumns());
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    var hdrRng = sh.getRange(1, 1, 1, N);
    var cur = hdrRng.getValues()[0];
    var need = headers.some(function(h,i){ return String(cur[i]||'').trim() !== h; });
    if (need) hdrRng.setValues([headers]);
    hdrRng
      .setFontFamily('Roboto').setFontSize(10).setFontWeight('bold').setFontColor('#ffffff')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    return sh;
  }
  function paintHeaderBlocks_OZ_(sh) {
    sh.getRange(1, 1, 1, 2).setBackground('#434343'); // A:B
    sh.getRange(1, 3, 1, 2).setBackground('#1c4587'); // C:D
    sh.getRange(1, 5, 1, 1).setBackground('#274e13'); // E
    sh.getRange(1, 6, 1, 3).setBackground('#6aa84f'); // F:H
    sh.getRange(1, 9, 1, 1).setBackground('#7f6000'); // I
    sh.getRange(1,10, 1, 1).setBackground('#990000'); // J
    sh.getRange(1,11, 1, 1).setBackground('#333333'); // K
    sh.getRange(1,12, 1, 2).setBackground('#5b0f00'); // L:M
    try {
      var a1Text = HEADERS[0]; // "[ OZ ] Кабинет"
      var tagEnd = a1Text.indexOf(']') + 1;
      var builder = SpreadsheetApp.newRichTextValue().setText(a1Text);
      if (tagEnd > 0) {
        builder.setTextStyle(0, tagEnd, SpreadsheetApp.newTextStyle().setForegroundColor('#ffff00').setBold(true).build());
        builder.setTextStyle(tagEnd, a1Text.length, SpreadsheetApp.newTextStyle().setForegroundColor('#ffffff').setBold(true).build());
      }
      sh.getRange('A1').setRichTextValue(builder.build());
    } catch(_) {}
  }
  function clearDataUnderHeader_(sh, N) {
    var last = sh.getLastRow();
    if (last >= 2) sh.getRange(2, 1, last - 1, N).clear({ contentsOnly: false });
  }
  function trimRowsAfter_(sh, lastRowToKeep) {
    var maxRows = sh.getMaxRows();
    var keep = Math.max(1, lastRowToKeep);
    if (maxRows > keep) sh.deleteRows(keep + 1, maxRows - keep);
  }
  function makeRowFillsFromCabinet_(rows, colorMap, cols) {
    if (!rows || !rows.length) return null;
    var fills = new Array(rows.length);
    for (var i = 0; i < rows.length; i++) {
      var cab = String(rows[i][0] || '').trim();
      var color = (colorMap && colorMap.get(cab)) || '#ffffff';
      var line = new Array(cols);
      for (var c = 0; c < cols; c++) line[c] = color;
      fills[i] = line;
    }
    return fills;
  }
  function applyPostStyles_(sh, startRow, numRows, cols) {
    if (numRows > 0) {
      sh.getRange(startRow, 1, numRows, cols)
        .setFontFamily('Roboto').setFontSize(10).setFontWeight('normal').setFontColor('#000000')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    }
    sh.autoResizeColumn(2);
    var w = sh.getColumnWidth(2);
    sh.setColumnWidth(2, w + 30);
  }
  // Читает «⚙️ Параметры»: A..H, фильтр D="OZON", H=TRUE
  function getSelectedCabinetsFromParams_OZ_() {
    var sh = ss.getSheetByName(PARAM_SHEET);
    if (!sh) return { list: [], total: 0 };
    var last = sh.getLastRow();
    if (last < 2) return { list: [], total: 0 };
    var rng = sh.getRange(2, 1, last - 1, 8).getValues(); // A..H
    var list = [], total = 0;
    for (var i = 0; i < rng.length; i++) {
      var name  = String(rng[i][0] || '').trim();              // A
      var platf = String(rng[i][3] || '').trim().toUpperCase();// D
      var on    = !!rng[i][7];                                 // H
      if (name && platf === 'OZON') {
        total++;
        if (on) list.push(name);
      }
    }
    return { list: list, total: total };
  }
}


/**
 * getREFRESHprices_OZ — обновляет J (Цена) на листе [OZ] Артикулы
 * и B (Цена) на «⛓️ Параллель», если активна OZ и совпадает кабинет.
 */
function getREFRESHprices_OZ() {
  var ss = SpreadsheetApp.getActive();
  var T0 = Date.now();
  var log = function (label, extra) {
    var ms = String(Date.now() - T0).padStart(7, ' ');
    console.log('[' + ms + ' ms] getREFRESHprices_OZ | ' + label + (extra ? ' | ' + extra : ''));
  };

  var SHEET = REF.SHEETS.ARTS_OZ;
  var sh = ss.getSheetByName(SHEET);
  if (!sh) { ss.toast('Лист не найден: ' + SHEET, 'OZON цены', 5); return; }
  var last = sh.getLastRow();
  if (last < 2) { ss.toast('Нет строк для обновления', 'OZON цены', 5); return; }

  var platTag = REF.getCurrentPlatform(); // 'OZ'|'WB'|null
var shCalc = ss.getSheetByName(REF.SHEETS.CALC);
var ctrlRange =
  (REF && typeof REF.getCabinetControlRange === 'function' && REF.getCabinetControlRange()) ||
  (shCalc ? shCalc.getRange(REF.CTRL_RANGE_A1) : null);
var currentCab = ctrlRange ? String(ctrlRange.getDisplayValue() || '').trim() : '';

  var normCab = REF.normCabinet;

  // Читаем A..K (A=Кабинет, B=Артикул, K=SKU опц)
  var rng = sh.getRange(2, 1, last - 1, 11).getValues();

  // group by cabinet
  var byCab = new Map(); // cab -> { offers:Set, rows:[{r, offer}] }
  for (var i = 0; i < rng.length; i++) {
    var cab   = String(rng[i][0] || '').trim();
    var offer = String(rng[i][1] || '').trim();
    if (!cab || !offer) continue;
    if (!byCab.has(cab)) byCab.set(cab, { offers: new Set(), rows: [] });
    var b = byCab.get(cab);
    b.offers.add(offer);
    b.rows.push({ r: i + 2, offer: offer });
  }

  var totalTouched = 0, totalUpdated = 0, totalParUpdated = 0;
  var successCabs = [];
  var CABs = Array.from(byCab.keys());

  for (var c = 0; c < CABs.length; c++) {
    var cabKey = CABs[c];
    ss.toast('Обновление цен: ' + cabKey, 'OZON', 3);
    var bundle = byCab.get(cabKey);
    var offers = Array.from(bundle.offers);

    var api = new OZONAPI(cabKey);
    var priceByOffer = {};
    var gotDirect = false;

    // 1) Быстрый метод (если доступен)
    try {
      if (typeof api.getPricesByOffers === 'function') {
        var CH = 100;
        for (var s = 0; s < offers.length; s += CH) {
          var chunk = offers.slice(s, s + CH);
          var part = api.getPricesByOffers(chunk) || {};
          for (var k in part) if (part.hasOwnProperty(k)) priceByOffer[k] = part[k];
        }
        gotDirect = Object.keys(priceByOffer).length > 0;
        log('getPricesByOffers OK', 'cab=' + cabKey + ', offers=' + offers.length + ', hit=' + Object.keys(priceByOffer).length);
      }
    } catch (e) {
      log('getPricesByOffers FAIL', 'cab=' + cabKey + ', err=' + ((e && e.message) || e));
    }

    // 2) Фолбэк — CSV-отчёт
    if (!gotDirect) {
      try {
        var reportId = api.makeArtsReport();
        Utilities.sleep(1000);
        var artsObj  = api.getArts(reportId) || {}; // { offerId: CSVcols }
        for (var j = 0; j < offers.length; j++) {
          var of = offers[j];
          var csv = artsObj[of];
          if (!csv) continue;
          var raw = String(csv['Текущая цена с учетом скидки, ₽'] || '').trim();
          var num = REF.toNumber(raw);
          if (!isNaN(num)) priceByOffer[of] = num;
        }
        log('CSV fallback OK', 'cab=' + cabKey + ', hit=' + Object.keys(priceByOffer).length);
      } catch (e2) {
        log('CSV fallback FAIL', 'cab=' + cabKey + ', err=' + ((e2 && e2.message) || e2));
      }
    }

    // 3) Обновляем J на [OZ] Артикулы
    var updates = [];
    var rows = bundle.rows;
    for (var t = 0; t < rows.length; t++) {
      var row = rows[t];
      totalTouched++;
      var price = priceByOffer[row.offer];
      if (price == null || price === '') continue;
      updates.push({ r: row.r, v: String(price).replace('.', ',') });
    }
    if (updates.length) {
      for (var u = 0; u < updates.length; u++) {
        sh.getRange(updates[u].r, 10).setValue(updates[u].v); // J
      }
      totalUpdated += updates.length;
      successCabs.push(cabKey);
    }

    // 4) «⛓️ Параллель» — только если активна OZ и кабинет совпадает с контролом
    if (platTag === 'OZ' && currentCab && normCab(cabKey) === normCab(currentCab)) {
      var shp = ss.getSheetByName('⛓️ Параллель');
      if (shp) {
        var lastP = shp.getLastRow();
        if (lastP >= 2) {
          var arts = shp.getRange(2, 1, lastP - 1, 1).getDisplayValues().map(function(r){ return String(r[0]||'').trim(); });
          var parUpdates = 0;
          for (var a = 0; a < arts.length; a++) {
            var art = arts[a]; if (!art) continue;
            var p = priceByOffer[art];
            if (p == null || p === '') continue;
            shp.getRange(a + 2, 2).setValue(String(p).replace('.', ',')); // B
            parUpdates++;
          }
          totalParUpdated += parUpdates;
          if (parUpdates) log('Parallel updated', 'cab=' + cabKey + ', rows=' + parUpdates);
        }
      }
    }
  }

  // ===== ЛОГ — только успешные кабинеты OZ
try { REF.logRun('Цены OZ', successCabs, 'OZON'); } catch(_){}

  ss.toast('OZON цены: найдено ' + totalUpdated + ' из ' + totalTouched + ' строк; Параллель: +' + totalParUpdated, 'Готово', 5);
  log('END', 'updated=' + totalUpdated + ', touched=' + totalTouched + ', parallel=' + totalParUpdated);
  triggerRebuildIfOZ_();                   // ← ДОБАВИТЬ
}

// === Перестроить калькулятор/параллель только если текущая платформа — OZON ===
function triggerRebuildIfOZ_() {
  try {
    var plat = (REF && typeof REF.getCurrentPlatform === 'function') ? REF.getCurrentPlatform() : null;
    if (plat !== 'OZ') return; // запускаем только на OZON

    var ss = SpreadsheetApp.getActive();
    var calcName = (REF && REF.SHEETS && REF.SHEETS.CALC) || '⚖️ Калькулятор';

    var shCalc = ss.getSheetByName(calcName);
    if (!shCalc) return;

    // именованный диапазон/контрол кабинета
    var ctrlRange =
      (REF && typeof REF.getCabinetControlRange === 'function' && REF.getCabinetControlRange()) ||
      shCalc.getRange(REF.CTRL_RANGE_A1);

    var currentCab = String(ctrlRange.getDisplayValue() || '').trim();
    if (!currentCab) return;

    // универсальный раннер этого проекта — без кулдаунов
    if (typeof runLayoutImmediate === 'function') {
      runLayoutImmediate(currentCab);
    } else {
      // запасной путь (на всякий случай)
      if (typeof layoutCalculator === 'function') layoutCalculator(currentCab, { plat: 'OZ' });
      if (typeof layoutParallelInline_ === 'function') layoutParallelInline_(currentCab, { plat: 'OZ' });
    }
  } catch (_) { /* тихо */ }
}

