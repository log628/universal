/** =========================================================
 * getREFRESH_WB — выгрузка Wildberries A:M (13 колонок) c фильтром «Раздел» (P:Q) и «Своя категория»
 * Колонки A..M:
 *  A Кабинет | B Артикул продавца | C Отзывы | D Рейтинг | E Категория | F FBO | G FBS | H RFBS
 *  I Объем | J Цена | K nmID | L Раздел | M Своя категория
 *
 * Фильтр «Раздел»:
 *  ⚙️ Параметры!P:P — список префиксов (со 2-й строки), Q:Q — «Своя категория».
 *  Для каждого vendorCode проверяем substring(3) (с 4-го символа) на startsWith(prefix).
 *  Если match — строка проходит, L="X", M=из Q. Иначе — строка НЕ выгружается.
 * ----------------------------------------------------------- */
function getREFRESH_WB() {
  var ss = SpreadsheetApp.getActive();
  var T0 = Date.now();
  var log = function (label, extra) {
    var ms = String(Date.now() - T0).padStart(7, ' ');
    console.log('[' + ms + ' ms] getREFRESH_WB | ' + label + (extra ? ' | ' + extra : ''));
  };

  log('START');

  var DST_SHEET   = REF.SHEETS.ARTS_WB;
  var PARAM_SHEET = REF.SHEETS.PARAMS;
  var HEADERS     = REF.getArtsHeaders('WB');   // A..M (13 колонок)
  var TOTAL_COLS  = HEADERS.length;             // 13

  // ===== 1) Префиксы «Раздел» и «Своя категория» (P:Q)
  var sections = (function readSections_() {
    var out = [];
    var sh = ss.getSheetByName(PARAM_SHEET);
    if (!sh) return out;
    var last = sh.getLastRow();
    if (last < 2) return out;
    // P (16) и Q (17)
    var vals = sh.getRange(2, 16, last - 1, 2).getDisplayValues();
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
    paintHeaderBlocks_WB_(sh0, HEADERS);
    trimRowsAfter_(sh0, 1);
    ss.toast('Нет префиксов в ⚙️ Параметры!P — выгрузка пуста', 'Готово', 5);
    log('END', 'no prefixes -> 0 rows');
    return;
  }

  // ===== 2) Выбранные WB-кабинеты (чекбоксы H) + карта токенов по ролям
  var pick = getSelectedCabinetsFromParams_WB_(); // { list, total }
  var selected = pick.list || [];
  var total    = pick.total || 0;
  log('PARAMS', 'WB total=' + total + ', selected=' + selected.length + ', cabinets=' + selected.length);

  var sh = ensureLayoutN_(ss, DST_SHEET, HEADERS, TOTAL_COLS);
  paintHeaderBlocks_WB_(sh, HEADERS);
  clearDataUnderHeader_(sh, TOTAL_COLS);

  if (!selected.length) {
    trimRowsAfter_(sh, 1);
    ss.toast('В ⚙️ Параметры ни один WB-кабинет не отмечен', 'Нет выбора', 5);
    log('END', 'no selected cabinets');
    return;
  }

  var colorMap = REF.readCabinetColorMap('WILDBERRIES'); // только WB-строки

  // Карта токенов по ролям: Map<cabinet, {prices[],content[],stats[],supplies[],any[]}>
  var roleMap = REF.buildWBTokenMapFromParams();

  // ===== 3) Фильтр «Раздел»
  function resolveSection_(vendorCode) {
    var s = String(vendorCode || '').toLowerCase();
    var tail = (s.length >= 4) ? s.substring(3) : ''; // с 4-го символа
    if (!tail) return null;
    for (var i = 0; i < sections.length; i++) {
      var p = sections[i];
      if (tail.indexOf(p.prefix) === 0) {
        return { section: 'X', ownCat: p.ownCat || '' };
      }
    }
    return null;
  }

  // Вызов API с подбором токена под роль
  function withToken_(cabName, role, callFn, tag) {
    var rec = roleMap.get(cabName) || { prices: [], content: [], stats: [], supplies: [], any: [] };
    var pools = [];
    if (role && rec[role] && rec[role].length) pools = rec[role].slice();
    else if (rec.any && rec.any.length) pools = rec.any.slice();

    var errs = [];
    for (var i = 0; i < pools.length; i++) {
      var tk = pools[i];
      try {
        var t0 = Date.now();
        var api = new WB(tk);
        var res = callFn(api, tk);
        log('API ' + tag + ' OK', 'cab=' + cabName + ', token#=' + (i + 1) + ', ' + (Date.now() - t0) + 'ms');
        return res;
      } catch (e) {
        var msg = (e && e.message) ? e.message : String(e);
        errs.push(msg);
        log('API ' + tag + ' FAIL', 'cab=' + cabName + ', token#=' + (i + 1) + ', err=' + msg);
      }
    }
    throw new Error('Кабинет «' + cabName + '»: ни один токен не сработал для ' + tag + '. Errs: ' + errs.join(' | '));
  }

  function buildRowsForCabinet(cabName) {
    ss.toast('Запрос данных кабинета ' + cabName, 'Выполнение', 3);

    // 1) Карточки (контент)
    var products = [];
    try {
      products = withToken_(cabName, 'content', function (api) { return api.getProducts(); }, 'getProducts') || [];
    } catch (e) {
      products = [];
      log('WARN getProducts', cabName + ': ' + ((e && e.message) || e));
    }
    if (!products.length) return [];

    // 2) Быстрые цены (list/goods/filter) — токен role=prices, фолбэк любой
    var pricesDict = {};
    try {
      var list = withToken_(cabName, 'prices', function (api) { return api.getPrices(); }, 'getPrices') || [];
      pricesDict = list.reduce(function (acc, it) { acc[it.nmID] = it; return acc; }, {});
    } catch (e) {
      pricesDict = {};
      log('WARN getPrices', cabName + ': ' + ((e && e.message) || e));
    }

    // 3) Комиссии (common-api) — любой рабочий токен
    var commissionsDict = {};
    try {
      var report = withToken_(cabName, 'any', function (api) { return api.getCategoryCommissions(); }, 'getCategoryCommissions') || [];
      for (var r = 0; r < report.length; r++) {
        var rec = report[r];
        var sid = (rec.subjectID != null ? rec.subjectID : rec.subjectId);
        if (!sid) continue;
        var FBO  = rec.kgvpMarketplace != null ? Number(rec.kgvpMarketplace) : '';
        var FBS  = rec.kgvpSupplier    != null ? Number(rec.kgvpSupplier)    : '';
        var RFBS = rec.kgvpBooking     != null ? Number(rec.kgvpBooking)     : '';
        commissionsDict[String(sid)] = { FBO: FBO, FBS: FBS, RFBS: RFBS };
      }
    } catch (e) {
      commissionsDict = {};
      log('WARN commissions', cabName + ': ' + ((e && e.message) || e));
    }

    // 4) Черновики строк + сбор nmID для RnR
    var drafts = [];
    var nmAll = [];
    for (var i = 0; i < products.length; i++) {
      var c = products[i];
      var nmID   = c.nmID || '';
      var vendor = c.vendorCode || c.vendor_code || '';
      var subjId = (c.subjectId != null ? c.subjectId : c.subjectID);
      var subjNm = c.subjectName || c.subject_name || '';

      // === ФИЛЬТР «Раздел»
      var sec = resolveSection_(vendor);
      if (!sec) continue;

      var vol = (function calcVolumeLiters_(card) {
        try {
          var d = card.dimensions || {};
          var L = Number(d.length) || 0;
          var W = Number(d.width)  || 0;
          var H = Number(d.height) || 0;
          if (!L || !W || !H) return '';
          var v = (L * W * H) / 1000;
          return (isFinite(v) ? String(v).replace('.', ',') : '');
        } catch(_) { return ''; }
      })(c);

      var price = (function pickPrice_(priceObj) {
        try {
          var p = priceObj && Array.isArray(priceObj.sizes) && priceObj.sizes[0] ? priceObj.sizes[0] : null;
          var v = (p && (p.price || p.clubDiscountedPrice || p.discountedPrice || p.basicSalePrice || p.basicPrice)) || '';
          return (v === '' ? '' : String(v).replace('.', ','));
        } catch(_) { return ''; }
      })(pricesDict[nmID]);

      drafts.push({
        cabName: cabName,
        vendor: vendor,
        nm: nmID,
        subjectId: subjId,
        subjectName: subjNm,
        vol: vol,
        price: price,
        section: sec.section,       // "X"
        ownCat: sec.ownCat || ''
      });
      if (nmID) nmAll.push(nmID);
    }

    // 5) RnR публичный
    var rnrDict = {};
    if (nmAll.length) {
      var page = 50;
      for (var idx = 0; idx < nmAll.length; idx += page) {
        var chunk = nmAll.slice(idx, idx + page);
        try {
          var url = 'https://card.wb.ru/cards/v2/detail?appType=1&curr=rub&dest=123589785&lang=ru&nm=' + chunk.join(';');
          var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
          var json = JSON.parse(resp.getContentText('UTF-8'));
          var arr = json && json.data && Array.isArray(json.data.products) ? json.data.products : [];
          for (var a = 0; a < arr.length; a++) rnrDict[arr[a].id] = arr[a];
        } catch (e) {
          log('WARN RnR batch fail', 'cab=' + cabName + ', offset=' + idx + ', err=' + ((e && e.message) || e));
        }
      }
    }

    // 6) Финальные строки (A..M = 13 колонок)
    var rows = [];
    for (var d = 0; d < drafts.length; d++) {
      var rec = drafts[d];
      var r   = rnrDict[rec.nm] || {};
      var rating   = (r.reviewRating != null ? r.reviewRating : (r.rating != null ? r.rating : '')) || '';
      var feedback = (r.feedbacks != null ? r.feedbacks : (r.feedbackCount != null ? r.feedbackCount : '')) || '';
      var cm = (rec.subjectId != null) ? (commissionsDict[String(rec.subjectId)] || {}) : {};
      rows.push([
        rec.cabName,           // A
        rec.vendor,            // B (Артикул продавца)
        (feedback || '0'),     // C
        (rating  || '0'),      // D
        (rec.subjectName || ''), // E
        (cm.FBO  == null ? '' : cm.FBO),   // F
        (cm.FBS  == null ? '' : cm.FBS),   // G
        (cm.RFBS == null ? '' : cm.RFBS),  // H
        rec.vol,               // I
        rec.price,             // J
        rec.nm,                // K (nmID)
        rec.section,           // L = "X"
        rec.ownCat             // M = «Своя категория»
      ]);
    }
    return rows;
  }

  var allRows = [];
  for (var i = 0; i < selected.length; i++) {
    var tCab = Date.now();
    var rows = [];
    try { rows = buildRowsForCabinet(selected[i]) || []; } catch (e) {
      log('WARN buildRowsForCabinet', selected[i] + ': ' + ((e && e.message) || e));
    }
    allRows = allRows.concat(rows);
    log('CAB done', selected[i] + ', rowsNow=' + allRows.length + ', took=' + (Date.now() - tCab) + 'ms');
  }

  // ===== 4) Запись
  clearDataUnderHeader_(sh, TOTAL_COLS);
  if (allRows.length) {
    sh.getRange(2, 1, allRows.length, TOTAL_COLS).setValues(allRows);

    // Заливки строк по цветам кабинетов (только WB цвета)
    var fills = makeRowFillsFromCabinet_(allRows, colorMap, TOTAL_COLS);
    if (fills) sh.getRange(2, 1, allRows.length, TOTAL_COLS).setBackgrounds(fills);

    applyPostStyles_(sh, 2, allRows.length, TOTAL_COLS);
  }

  // ===== 5) Обрезка хвоста
  trimRowsAfter_(sh, (allRows.length || 0) + 1);

  ss.toast('Готово! Обновлено ' + (allRows.length || 0) + ' строк (WB)', DST_SHEET, 5);
  log('END', 'totalRows=' + (allRows.length || 0));

  /* ========================= ХЕЛПЕРЫ ========================= */

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

  function paintHeaderBlocks_WB_(sh, HEADERS_) {
    sh.getRange(1,  1, 1, 2).setBackground('#434343'); // A:B — Кабинет, Артикул
    sh.getRange(1,  3, 1, 2).setBackground('#1c4587'); // C:D — Отзывы, Рейтинг
    sh.getRange(1,  5, 1, 1).setBackground('#274e13'); // E   — Категория
    sh.getRange(1,  6, 1, 3).setBackground('#6aa84f'); // F:H — FBO,FBS,RFBS
    sh.getRange(1,  9, 1, 1).setBackground('#7f6000'); // I   — Объем
    sh.getRange(1, 10, 1, 1).setBackground('#990000'); // J   — Цена
    sh.getRange(1, 11, 1, 1).setBackground('#333333'); // K   — nmID
    sh.getRange(1, 12, 1, 2).setBackground('#5b0f00'); // L:M — Раздел, Своя категория

    // Жёлтый тег "[ WB ]" в A1
    try {
      var a1Text = HEADERS_[0]; // "[ WB ] Кабинет"
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
      var cab = String(rows[i][0] || '').replace(/\u00A0/g, ' ').trim();
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

  // Читает «⚙️ Параметры»: A=Кабинет, D=Площадка/Тип, H=чекбокс (TRUE/FALSE), фильтр: D in {"WILDBERRIES","WB"}
  function getSelectedCabinetsFromParams_WB_() {
    var sh = ss.getSheetByName(PARAM_SHEET);
    if (!sh) return { list: [], total: 0 };
    var last = sh.getLastRow();
    if (last < 2) return { list: [], total: 0 };
    var rng = sh.getRange(2, 1, last - 1, 8).getDisplayValues(); // A..H
    var list = [], total = 0;
    for (var i = 0; i < rng.length; i++) {
      var name  = String(rng[i][0] || '').trim();               // A
      var platf = String(rng[i][3] || '').trim().toUpperCase(); // D
      var on    = !!rng[i][7];                                  // H
      if (name && (platf === 'WILDBERRIES' || platf === 'WB')) {
        total++;
        if (on) list.push(name);
      }
    }
    return { list: list, total: total };
  }
}
