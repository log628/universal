/** getREFRESH_OZ — выгрузка A:M с фильтром «Раздел» (P:Q) и «Своя категория»
 *  Колонки: A..M = 13:
 *   A Кабинет | B Артикул | C Отзывы | D Рейтинг | E Категория | F FBO | G FBS | H RFBS
 *   I Объем | J Цена | K SKU | L Раздел | M Своя категория
 *
 *  Фильтр «Раздел»:
 *   ⚙️ Параметры!P:P — список префиксов (со 2-й строки), Q:Q — «Своя категория» для того же префикса.
 *   Для каждого offer_id проверяем substring(3) (с 4-го символа) на startsWith(prefix).
 *   Если match — строка проходит, L="X", M=из Q. Иначе — строка НЕ выгружается.
 */
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

  // ====== Константы/шапка (REF уже без «Наименование») =====
  var TOTAL_COLS = 13; // A..M
  var HEADERS = REF.getArtsHeaders('OZ'); // уже 13 колонок, A:M

  // ====== 1) Префиксы «Раздел» и «Своя категория» (P:Q)
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
    paintHeaderBlocks_OZ_(sh0); // включая #5b0f00 для L:M
    trimRowsAfter_(sh0, 1);
    ss.toast('Нет префиксов в ⚙️ Параметры!P — выгрузка пуста', 'Готово', 5);
    log('END', 'no prefixes -> 0 rows');
    return;
  }

  // ====== 2) Выбранные кабинеты OZON (чекбоксы H)
  var pick = getSelectedCabinetsFromParams_OZ_(); // { list, total }
  var selected = pick.list || [];
  var total    = pick.total || 0;

  var sh = ensureLayoutN_(ss, DST_SHEET, HEADERS, TOTAL_COLS);
  paintHeaderBlocks_OZ_(sh);
  clearDataUnderHeader_(sh, TOTAL_COLS);

  if (!selected.length) {
    trimRowsAfter_(sh, 1);
    ss.toast('В ⚙️ Параметры ни один OZON-кабинет не отмечен', 'Нет выбора', 5);
    log('END', 'no selected cabinets');
    return;
  }

  var colorMap = REF.readCabinetColorMap('OZON');

  // ====== 3) Сборка данных по кабинетам
  function buildRowsForCabinet(cabKey) {
    ss.toast('Запрос данных кабинета ' + cabKey, 'Выполнение', 3);
    var api = new OZONAPI(cabKey);

    var reportId = api.makeArtsReport();
    Utilities.sleep(1000);
    var artsObj  = api.getArts(reportId) || {};         // { offerId: CSVcols }
    var offers   = Object.keys(artsObj);
    if (!offers.length) return [];

    var details = api.getArtsDetails(offers) || [];     // items[]
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
    var byCategoryId = cats.byCategoryId || {};
    var byTypeId     = cats.byTypeId || {};

    function pickCategoryTitle_(det) {
      var cid = det && (det.category_id || det.description_category_id);
      var tid = det && det.type_id;
      if (cid != null && byCategoryId[String(cid)]) return byCategoryId[String(cid)];
      if (tid != null && byTypeId[String(tid)])     return byTypeId[String(tid)];
      return '';
    }
    function pickCommissions_(det) {
      var fboPct = '', fbsPct = '', rfbsPct = '';
      var arr = det && Array.isArray(det.commissions) ? det.commissions : [];
      for (var i = 0; i < arr.length; i++) {
        var c = arr[i];
        if (c.sale_schema === 'FBO')  fboPct = c.percent || '';
        if (c.sale_schema === 'FBS')  fbsPct = c.percent || '';
        if (c.sale_schema === 'RFBS') rfbsPct = c.percent || '';
      }
      return { fboPct: fboPct, fbsPct: fbsPct, rfbsPct: rfbsPct };
    }
    function pickSku_(det, offer) {
      return (det && (det.sku || det.sku_id || det.id)) || skuFallbackMap[offer] || '';
    }
    function calcVolumeLitersFromCsv_(csv) {
      var s = String(csv['Объем товара, л'] || '').replace(/\./g, ',');
      return s;
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
      var offer = offers[i];                 // B
      var csv   = artsObj[offer] || {};
      var det   = detMap[offer] || {};

      // === ФИЛЬТР «Раздел»
      var sec = resolveSection_(offer);
      if (!sec) continue;

      var cabinet = cabKey;                 // A
      var reviews = csv['Отзывы'] || '';    // C
      var rating  = String(csv['Рейтинг'] || '').replace(/\./g, ','); // D

      var categoryTitle = pickCategoryTitle_(det); // E
      var cm            = pickCommissions_(det);   // F,G,H
      var volStr        = calcVolumeLitersFromCsv_(csv); // I
      var priceStr      = pickPriceFromCsv_(csv);        // J
      var sku           = pickSku_(det, offer);          // K
      var sectionLabel  = sec.section;                   // L = "X"
      var ownCategory   = sec.ownCat;                    // M

      rows.push([
        cabinet,     // A
        offer,       // B
        reviews,     // C
        rating,      // D
        categoryTitle, // E
        cm.fboPct,   // F
        cm.fbsPct,   // G
        cm.rfbsPct,  // H
        volStr,      // I
        priceStr,    // J
        sku,         // K
        sectionLabel,// L
        ownCategory  // M
      ]);
    }

    return rows;
  }

  // ====== 4) Режимы: [ВСЕ] / мультивыбор
  var allRows = [];
  for (var i = 0; i < selected.length; i++) {
    var tCab = Date.now();
    var rows = [];
    try {
      rows = buildRowsForCabinet(selected[i]) || [];
    } catch (e) {
      log('WARN buildRowsForCabinet', selected[i] + ': ' + ((e && e.message) || e));
    }
    allRows = allRows.concat(rows);
    log('CAB done', selected[i] + ', rowsNow=' + allRows.length + ', took=' + (Date.now() - tCab) + 'ms');
  }

  // ====== 5) Запись
  clearDataUnderHeader_(sh, TOTAL_COLS);
  if (allRows.length) {
    sh.getRange(2, 1, allRows.length, TOTAL_COLS).setValues(allRows);

    // Заливки строк по цветам кабинетов
    var fills = makeRowFillsFromCabinet_(allRows, colorMap, TOTAL_COLS);
    if (fills) sh.getRange(2, 1, allRows.length, TOTAL_COLS).setBackgrounds(fills);

    // Пост-стили
    applyPostStyles_(sh, 2, allRows.length, TOTAL_COLS);
  }

  // ====== 6) Обрезка хвоста
  trimRowsAfter_(sh, (allRows.length || 0) + 1);

  ss.toast('Готово! Обновлено ' + (allRows.length || 0) + ' строк (OZON)', DST_SHEET, 5);
  log('END', 'totalRows=' + (allRows.length || 0));

  /* ==================== ЛОКАЛЬНЫЕ ХЕЛПЕРЫ ==================== */

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

    // Стиль шапки
    hdrRng
      .setFontFamily('Roboto').setFontSize(10).setFontWeight('bold').setFontColor('#ffffff')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');

    return sh;
  }

  function paintHeaderBlocks_OZ_(sh) {
    // Группы под 13 колонок; L:M = Раздел/Своя категория
    sh.getRange(1, 1, 1, 2).setBackground('#434343'); // A:B — Кабинет, Артикул
    sh.getRange(1, 3, 1, 2).setBackground('#1c4587'); // C:D — Отзывы, Рейтинг
    sh.getRange(1, 5, 1, 1).setBackground('#274e13'); // E   — Категория
    sh.getRange(1, 6, 1, 3).setBackground('#6aa84f'); // F:H — FBO,FBS,RFBS
    sh.getRange(1, 9, 1, 1).setBackground('#7f6000'); // I   — Объем
    sh.getRange(1,10, 1, 1).setBackground('#990000'); // J   — Цена
    sh.getRange(1,11, 1, 1).setBackground('#333333'); // K   — SKU
    sh.getRange(1,12, 1, 2).setBackground('#5b0f00'); // L:M — Раздел, Своя категория

    // Жёлтый тег "[ OZ ]" в A1
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

  // Читает «⚙️ Параметры»: A=Кабинет, D=Площадка/Тип, H=чекбокс (TRUE/FALSE), фильтр: D="OZON"
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
