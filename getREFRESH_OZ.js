/** ===================== getREFRESH_OZ (SKU + Баркод, алиасы заголовков) ===================== */
function getREFRESH_OZ() {
  var ss = SpreadsheetApp.getActive();
  var T0 = Date.now();
  var log = function (label, extra) {
    var ms = String(Date.now() - T0).padStart(7, ' ');
    console.log('[' + ms + ' ms] getREFRESH_OZ | ' + label + (extra ? ' | ' + extra : ''));
  };

  // === Канонические имена и алиасы под шапку REF ===
  // Канон соответствует REF.ARTS_HEADERS_BASE (+добавим "Баркод" после SKU на лету)
  var CANON = {
    CAB: '[ OZ ] Кабинет',
    OFFER: 'Артикул',
    REVIEWS: 'Отзывы',
    RATING: 'Рейтинг',
    CATEGORY: 'Категория',
    FBO: 'FBO',
    FBS: 'FBS',
    RFBS: 'RFBS',
    VOLUME: 'Объем',
    PRICE: 'Цена',
    SKU: 'SKU',
    BARCODE: 'Баркод',
    SECTION: 'Раздел',
    OWN: 'Своя категория'
  };

  // Что бы ни приехало из CSV/деталей — приведём к канону
  var ALIAS = new Map([
    // Комиссии (любые вариации с %/запятыми/пробелами)
    [CANON.FBO, ['FBO', 'Комиссия FBO %', 'Комиссия FBO, %', 'FBO %', 'FBO, %']],
    [CANON.FBS, ['FBS', 'Комиссия FBS %', 'Комиссия FBS, %', 'FBS %', 'FBS, %']],
    [CANON.RFBS, ['RFBS','Комиссия RFBS %','Комиссия RFBS, %','RFBS %','RFBS, %']],
    // Объём
    [CANON.VOLUME, ['Объем', 'Объём', 'Объем товара, л', 'Объём товара, л', 'Объем, л', 'Объём, л', 'Объем товара']],
    // Цена
    [CANON.PRICE, ['Цена', 'Цена, ₽', 'Текущая цена с учетом скидки, ₽', 'Текущая цена с учетом скидки', 'Цена со скидкой']],
    // Категория
    [CANON.CATEGORY, ['Категория', 'Категория комиссии']]
  ]);

  function putByHeaderName(line, HEADERS, canonKey, value) {
    var norm = function(s){ return String(s||'').trim(); };
    var idx = HEADERS.findIndex(function(h){ return norm(h) === norm(canonKey); });
    if (idx < 0 && ALIAS.has(canonKey)) {
      var variants = ALIAS.get(canonKey);
      idx = HEADERS.findIndex(function(h){ return variants.some(function(a){ return norm(a) === norm(h); }); });
    }
    if (idx >= 0) line[idx] = value;
  }

  function normOffer_(s){ return String(s||'').trim().replace(/'/g,''); }

  function pickCsv_(csv, variants) {
    // 1) Точное совпадение по ключу
    for (var i = 0; i < variants.length; i++) {
      var k = variants[i];
      if (csv.hasOwnProperty(k) && csv[k] !== '') return csv[k];
    }
    // 2) Мягкий поиск по нормализованному названию
    var keys = Object.keys(csv || {});
    for (var j = 0; j < keys.length; j++) {
      var key = keys[j];
      var normKey = String(key).toLowerCase().replace(/\s+/g,' ');
      for (var v = 0; v < variants.length; v++) {
        var pat = String(variants[v]).toLowerCase().replace(/\s+/g,' ');
        if (normKey.indexOf(pat) >= 0) return csv[key];
      }
    }
    return '';
  }

  function toNumberSafe_(v) {
    if (typeof REF !== 'undefined' && typeof REF.toNumber === 'function') return REF.toNumber(v);
    var s = String(v == null ? '' : v).replace(/\s/g,'').replace(',', '.');
    var n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }

  log('START');

  var DST_SHEET   = REF.SHEETS.ARTS_OZ;
  var PARAM_SHEET = REF.SHEETS.PARAMS;

  // --- Заголовки: берём из REF и если нет "Баркод", вставим после SKU
  var HEADERS_ORIG = REF.getArtsHeaders('OZ'); // база из Refs: 'Кабинет','Артикул','Отзывы','Рейтинг','Категория','FBO','FBS','RFBS','Объем','Цена','SKU','Раздел','Своя категория'
  var HEADERS = (function ensureHeadersWithBarcode_(hdr) {
    var H = hdr.slice();
    var norm = function(s){ return String(s||'').trim().toLowerCase(); };
    var iSKU = H.findIndex(function(x){ return norm(x)==='sku'; });
    var iBar = H.findIndex(function(x){ return ['баркод','barcode','штрихкод'].indexOf(norm(x))!==-1; });
    if (iSKU >= 0 && iBar < 0) H.splice(iSKU + 1, 0, CANON.BARCODE); // сразу после SKU
    return H;
  })(HEADERS_ORIG);

  var TOTAL_COLS = HEADERS.length;

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
    paintHeaderBlocks_OZ_(sh0, HEADERS);
    trimRowsAfter_(sh0, 1);
    ss.toast('Нет префиксов в ⚙️ Параметры!P — выгрузка пуста', 'Готово', 5);
    log('END', 'no prefixes -> 0 rows');
    try { REF.logRun('Артикулы OZ', [], 'OZON'); } catch(_){}
    try { REF.logRun('Цены OZ',      [], 'OZON'); } catch(_){}
    return;
  }

  // ===== 2) Выбранные OZ-кабинеты (чекбокс H)
  var pick = getSelectedCabinetsFromParams_OZ_(); // { list, total }
  var selected = pick.list || [];

  var sh = ensureLayoutN_(ss, DST_SHEET, HEADERS, TOTAL_COLS);
  paintHeaderBlocks_OZ_(sh, HEADERS);
  clearDataUnderHeader_(sh, TOTAL_COLS);

  if (!selected.length) {
    trimRowsAfter_(sh, 1);
    ss.toast('В ⚙️ Параметры ни один OZON-кабинет не отмечен', 'Нет выбора', 5);
    log('END', 'no selected cabinets');
    try { REF.logRun('Артикулы OZ', [], 'OZON'); } catch(_){}
    try { REF.logRun('Цены OZ',      [], 'OZON'); } catch(_){}
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
      var offer = normOffer_(it && (it.offer_id || it.offer || ''));
      if (offer) acc[offer] = it;
      return acc;
    }, {});

    var missingForSku = offers.filter(function (of) {
      var d = detMap[normOffer_(of)];
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
      var arr = [];
      if (det && Array.isArray(det.commissions)) arr = det.commissions;
      else if (det && det.financial && Array.isArray(det.financial.commissions)) arr = det.financial.commissions;

      for (var i = 0; i < arr.length; i++) {
        var c = arr[i] || {};
        var schema = String(c.sale_schema || c.schema || '').toUpperCase();
        var val = (c.percent != null) ? toNumberSafe_(c.percent) : (c.value != null ? toNumberSafe_(c.value) : NaN);
        if (!isNaN(val)) {
          if (schema === 'FBO')  fboPct  = val;
          if (schema === 'FBS')  fbsPct  = val;
          if (schema === 'RFBS') rfbsPct = val;
        }
      }
      return { fboPct: fboPct, fbsPct: fbsPct, rfbsPct: rfbsPct };
    }

    function pickSku_(det, offer) {
      return (det && (det.sku || det.sku_id || det.id)) || skuFallbackMap[offer] || '';
    }

    function pickBarcode_(det, csv) {
      var v = '';
      if (det) {
        if (Array.isArray(det.barcodes) && det.barcodes.length) v = String(det.barcodes[0] || '');
        else if (det.barcode) v = String(det.barcode);
      }
      if (!v && csv) {
        var raw = pickCsv_(csv, ['Штрихкод','Штрихкод товара','Штрихкоды','EAN','EAN-13','ean','ean13','Barcode','Баркод']);
        if (raw) v = String(raw).split(/[;,|\s]+/)[0];
      }
      return v || '';
    }

    function calcVolumeLitersFromCsv_(csv) {
      var raw = pickCsv_(csv, ALIAS.get(CANON.VOLUME) || [CANON.VOLUME]);
      return String(raw || '').replace(/\./g, ',');
    }

    function pickPriceFromCsv_(csv) {
      var raw = pickCsv_(csv, ALIAS.get(CANON.PRICE) || [CANON.PRICE]);
      var num = toNumberSafe_(raw);
      return num === 0 && !raw ? '' : String(num).replace('.', ',');
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
      var det   = detMap[normOffer_(offer)] || {};
      var sec   = resolveSection_(offer);
      if (!sec) continue;

      var cm = pickCommissions_(det);
      var ct = pickCategoryAndType_(det);
      var catCell = ct.category || '';
      if (ct.typeName) catCell += ' | ' + ct.typeName;

      // Готовим значения
      var valCab     = String(cabKey);
      var valOffer   = offer;
      var valReviews = csv['Отзывы'] || '';
      var valRating  = String(csv['Рейтинг'] || '').replace(/\./g, ',');
      var valCat     = catCell;
      var valFBO     = cm.fboPct;
      var valFBS     = cm.fbsPct;
      var valRFBS    = cm.rfbsPct;
      var valVolume  = calcVolumeLitersFromCsv_(csv);
      var valPrice   = pickPriceFromCsv_(csv);
      var valSku     = pickSku_(det, offer);
      var valBarcode = pickBarcode_(det, csv);
      var valSection = 'X';
      var valOwn     = sec.ownCat || '';

      // Строго укладываем по HEADERS с учётом алиасов
      var line = new Array(HEADERS.length).fill('');

      putByHeaderName(line, HEADERS, CANON.CAB,     valCab);
      putByHeaderName(line, HEADERS, CANON.OFFER,   valOffer);
      putByHeaderName(line, HEADERS, CANON.REVIEWS, valReviews);
      putByHeaderName(line, HEADERS, CANON.RATING,  valRating);
      putByHeaderName(line, HEADERS, CANON.CATEGORY,valCat);

      putByHeaderName(line, HEADERS, CANON.FBO,     valFBO);
      putByHeaderName(line, HEADERS, CANON.FBS,     valFBS);
      putByHeaderName(line, HEADERS, CANON.RFBS,    valRFBS);

      putByHeaderName(line, HEADERS, CANON.VOLUME,  valVolume);
      putByHeaderName(line, HEADERS, CANON.PRICE,   valPrice);

      putByHeaderName(line, HEADERS, CANON.SKU,     valSku);
      putByHeaderName(line, HEADERS, CANON.BARCODE, valBarcode);
      putByHeaderName(line, HEADERS, CANON.SECTION, valSection);
      putByHeaderName(line, HEADERS, CANON.OWN,     valOwn);

      rows.push(line);
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
    var fills = makeRowFillsFromCabinet_(allRows, REF.readCabinetColorMap('OZON'), TOTAL_COLS);
    if (fills) sh.getRange(2, 1, allRows.length, TOTAL_COLS).setBackgrounds(fills);
    applyPostStyles_(sh, 2, allRows.length, TOTAL_COLS);
  }
  trimRowsAfter_(sh, (allRows.length || 0) + 1);

  ss.toast('Готово! Обновлено ' + (allRows.length || 0) + ' строк (OZON)', DST_SHEET, 5);
  log('END', 'totalRows=' + (allRows.length || 0));

  // ===== 5) Штампы
  try { REF.logRun('Артикулы OZ', successCabs, 'OZON'); } catch(_){}
  try { REF.logRun('Цены OZ',      successCabs, 'OZON'); } catch(_){}

  /* ==================== локальные хелперы/оформление ==================== */
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
  function paintHeaderBlocks_OZ_(sh, headers) {
    var H = headers.map(function(s){return String(s).trim().toLowerCase();});
    var cSku    = H.indexOf('sku')>=0 ? (H.indexOf('sku')+1) : 11;
    var cBarcode= H.indexOf('баркод')>=0 ? (H.indexOf('баркод')+1) : (cSku+1);
    var cSection= H.indexOf('раздел')>=0 ? (H.indexOf('раздел')+1) : (cBarcode+1);
    var cOwn    = H.indexOf('своя категория')>=0 ? (H.indexOf('своя категория')+1) : (cSection+1);

    sh.getRange(1, 1, 1, 2).setBackground('#434343'); // A:B
    sh.getRange(1, 3, 1, 2).setBackground('#1c4587'); // C:D
    sh.getRange(1, 5, 1, 1).setBackground('#274e13'); // E
    sh.getRange(1, 6, 1, 3).setBackground('#6aa84f'); // F:H
    sh.getRange(1, 9, 1, 1).setBackground('#7f6000'); // I
    sh.getRange(1,10, 1, 1).setBackground('#990000'); // J
    sh.getRange(1, cSku, 1, Math.max(1, cBarcode - cSku + 1)).setBackground('#333333'); // SKU+Баркод
    sh.getRange(1, cSection, 1, Math.max(1, (cOwn - cSection + 1))).setBackground('#5b0f00'); // Раздел+Своя

    try {
      var a1Text = headers[0]; // "[ OZ ] Кабинет"
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
 * getREFRESHprices_OZ — обновляет колонку «Цена» на листе [OZ] Артикулы
 * и колонку B (Цена) на листе «⛓️ Параллель», если активна OZ и выбранный кабинет совпадает.
 * Устойчив к изменениям шапки (ищет «Цена» по имени/алиасам, а не по индексу J).
 */
function getREFRESHprices_OZ() {
  var ss = SpreadsheetApp.getActive();
  var T0 = Date.now();
  var log = function (label, extra) {
    var ms = String(Date.now() - T0).padStart(7, ' ');
    console.log('[' + ms + ' ms] getREFRESHprices_OZ | ' + label + (extra ? ' | ' + extra : ''));
  };

  var SHEET = REF.SHEETS.ARTS_OZ; // "[OZ] Артикулы"
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    ss.toast('Лист не найден: ' + SHEET, 'OZON цены', 5);
    return;
  }
  var last = sh.getLastRow();
  if (last < 2) {
    ss.toast('Нет строк для обновления', 'OZON цены', 5);
    return;
  }

  // --- Алиасы для «Цена» (на случай разных формулировок)
  var PRICE_ALIASES = ['Цена', 'Цена, ₽', 'Текущая цена с учетом скидки, ₽', 'Текущая цена с учетом скидки', 'Цена со скидкой'];

  function findColumnByHeader_(headers, names) {
    var norm = function (s) { return String(s || '').trim().toLowerCase(); };
    for (var i = 0; i < headers.length; i++) {
      var h = norm(headers[i]);
      for (var j = 0; j < names.length; j++) {
        if (h === norm(names[j])) return i + 1; // 1-based
      }
    }
    return 0;
  }

  // --- Определим индексы колонок по шапке
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  var colCab   = findColumnByHeader_(hdr, ['[ oz ] кабинет', '[oz] кабинет', 'кабинет']); // обычно A
  var colOffer = findColumnByHeader_(hdr, ['артикул']);                                    // обычно B
  var colPrice = findColumnByHeader_(hdr, PRICE_ALIASES);                                  // обычно J
  if (!colCab || !colOffer || !colPrice) {
    ss.toast('Не найдены нужные колонки (Кабинет/Артикул/Цена) в шапке', 'OZON цены', 8);
    return;
  }

  // --- Платформа/выбранный кабинет для «⛓️ Параллель»
  var platTag = REF.getCurrentPlatform(); // 'OZ'|'WB'|null
  var shCalc = ss.getSheetByName(REF.SHEETS.CALC);
  var ctrlRange = (REF && typeof REF.getCabinetControlRange === 'function' && REF.getCabinetControlRange())
               || (shCalc ? shCalc.getRange(REF.CTRL_RANGE_A1) : null);
  var currentCab = ctrlRange ? String(ctrlRange.getDisplayValue() || '').trim() : '';
  var normCab = REF.normCabinet;

  // === Считываем строки (только A..max для скорости)
  var vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getDisplayValues();

  // --- Группируем по кабинету и собираем список артикула/строки
  var byCab = new Map(); // cab -> { offers:Set, rows:[{r, offer}] }
  for (var r = 0; r < vals.length; r++) {
    var cab = String(vals[r][colCab - 1] || '').trim();
    var offer = String(vals[r][colOffer - 1] || '').trim();
    if (!cab || !offer) continue;
    if (!byCab.has(cab)) byCab.set(cab, { offers: new Set(), rows: [] });
    var b = byCab.get(cab);
    b.offers.add(offer);
    b.rows.push({ r: r + 2, offer: offer });
  }

  var totalTouched = 0, totalUpdated = 0, totalParUpdated = 0;
  var successCabs = [];
  var CABs = Array.from(byCab.keys());

  // --- Хелпер: надёжный парс числа
  function toNumberSafe_(v) {
    if (typeof REF !== 'undefined' && typeof REF.toNumber === 'function') return REF.toNumber(v);
    var s = String(v == null ? '' : v).replace(/\s/g, '').replace(',', '.');
    var n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }

  // --- Хелпер: взять цену из CSV записи по множеству возможных ключей
  function pickCsvPrice_(csv) {
    if (!csv) return '';
    // 1) точные ключи
    for (var i = 0; i < PRICE_ALIASES.length; i++) {
      var k = PRICE_ALIASES[i];
      if (csv.hasOwnProperty(k) && csv[k] !== '') return csv[k];
    }
    // 2) гибкий поиск по нормализованному названию
    var keys = Object.keys(csv || {});
    for (var j = 0; j < keys.length; j++) {
      var key = keys[j];
      var normKey = String(key).toLowerCase().replace(/\s+/g, ' ');
      for (var v = 0; v < PRICE_ALIASES.length; v++) {
        var pat = String(PRICE_ALIASES[v]).toLowerCase().replace(/\s+/g, ' ');
        if (normKey.indexOf(pat) >= 0) return csv[key];
      }
    }
    return '';
  }

  for (var c = 0; c < CABs.length; c++) {
    var cabKey = CABs[c];
    ss.toast('Обновление цен: ' + cabKey, 'OZON', 3);

    var bundle = byCab.get(cabKey);
    var offers = Array.from(bundle.offers);
    var api = new OZONAPI(cabKey);

    var priceByOffer = {};
    var gotDirect = false;

    // 1) Быстрый метод (если есть реализация api.getPricesByOffers)
    try {
      if (typeof api.getPricesByOffers === 'function') {
        var CH = 100;
        for (var s = 0; s < offers.length; s += CH) {
          var chunk = offers.slice(s, s + CH);
          var part = api.getPricesByOffers(chunk) || {}; // ожидаем { offerId: number }
          for (var k in part) if (part.hasOwnProperty(k)) priceByOffer[k] = part[k];
        }
        gotDirect = Object.keys(priceByOffer).length > 0;
        log('getPricesByOffers OK', 'cab=' + cabKey + ', offers=' + offers.length + ', hit=' + Object.keys(priceByOffer).length);
      }
    } catch (e) {
      log('getPricesByOffers FAIL', 'cab=' + cabKey + ', err=' + ((e && e.message) || e));
    }

    // 2) Фолбэк — CSV-отчёт об артикулах и чтение цены из CSV полей
    if (!gotDirect) {
      try {
        var reportId = api.makeArtsReport();
        Utilities.sleep(1000);
        var artsObj = api.getArts(reportId) || {}; // { offerId: CSVcols }
        for (var j = 0; j < offers.length; j++) {
          var of = offers[j];
          var csv = artsObj[of];
          if (!csv) continue;
          var raw = String(pickCsvPrice_(csv) || '').trim();
          var num = toNumberSafe_(raw);
          if (!isNaN(num)) priceByOffer[of] = num;
        }
        log('CSV fallback OK', 'cab=' + cabKey + ', hit=' + Object.keys(priceByOffer).length);
      } catch (e2) {
        log('CSV fallback FAIL', 'cab=' + cabKey + ', err=' + ((e2 && e2.message) || e2));
      }
    }

    // 3) Обновляем колонку «Цена» на [OZ] Артикулы (по найденному индексу colPrice)
    var rows = bundle.rows;
    var updatesCount = 0;
    for (var t = 0; t < rows.length; t++) {
      var row = rows[t];
      totalTouched++;
      var price = priceByOffer[row.offer];
      if (price == null || price === '') continue;
      sh.getRange(row.r, colPrice).setValue(String(price).replace('.', ',')); // безопасно по индексу колонки «Цена»
      updatesCount++;
    }
    if (updatesCount) {
      totalUpdated += updatesCount;
      successCabs.push(cabKey);
    }

    // 4) «⛓️ Параллель» — только если активна OZ и кабинет совпадает с контролом
    if (platTag === 'OZ' && currentCab && normCab(cabKey) === normCab(currentCab)) {
      var shp = ss.getSheetByName('⛓️ Параллель');
      if (shp) {
        var lastP = shp.getLastRow();
        if (lastP >= 2) {
          var arts = shp.getRange(2, 1, lastP - 1, 1).getDisplayValues().map(function (r) { return String(r[0] || '').trim(); });
          var parUpdates = 0;
          for (var a = 0; a < arts.length; a++) {
            var art = arts[a];
            if (!art) continue;
            var p = priceByOffer[art];
            if (p == null || p === '') continue;
            shp.getRange(a + 2, 2).setValue(String(p).replace('.', ',')); // B = Цена
            parUpdates++;
          }
          totalParUpdated += parUpdates;
          if (parUpdates) log('Parallel updated', 'cab=' + cabKey + ', rows=' + parUpdates);
        }
      }
    }
  }

  // Лог шапки в ⚙️ Параметры (ставит штамп T и список кабинетов в U)
  try { REF.logRun('Цены OZ', successCabs, 'OZON'); } catch (_) {}

  ss.toast('OZON цены: найдено ' + totalUpdated + ' из ' + totalTouched + ' строк; Параллель: +' + totalParUpdated, 'Готово', 5);
  log('END', 'updated=' + totalUpdated + ', touched=' + totalTouched + ', parallel=' + totalParUpdated);
}

