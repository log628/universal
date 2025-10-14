/** =========================================================
 * Универсальная отправка цен из «⚖️ Калькулятор» в включённую площадку (⚙️ Параметры!I2)
 * + Подробный WB-лог со статусами/ошибками/карантином в лист «🛠 Тех. лог» (с H)
 * ========================================================= */
function sendPricesFromCalculatorFast() {
  var T0 = Date.now();
  function log(label, extra) {
    var ms = String(Date.now() - T0).padStart(6, ' ');
    console.log('[' + ms + ' ms][sendPricesFromCalculatorFast] ' + label + (extra ? ' | ' + extra : ''));
  }

  try {
    var ss = SpreadsheetApp.getActive();
    var shCalc = ss.getSheetByName('⚖️ Калькулятор');
    var shPar  = ss.getSheetByName('⚙️ Параметры');
    if (!shCalc || !shPar) throw new Error('Не найден «⚖️ Калькулятор» или «⚙️ Параметры»');

    // Площадка строго из I2
// === Площадка ТОЛЬКО из именованного muff_mp (REF.NAMED.MP_CTRL)
var rMP   = (REF && REF.NAMED && REF.NAMED.MP_CTRL) ? ss.getRangeByName(REF.NAMED.MP_CTRL) : null;
var mpRaw = rMP ? String(rMP.getDisplayValue() || '').trim() : '';
var PLAT  = (function (s) {
  if (REF && typeof REF.platformCanon === 'function') {
    var c = REF.platformCanon(s);            // 'OZ' | 'WB' | null
    if (c) return c;
  }
  if (/wb|wildberries/i.test(s)) return 'WB'; // мягкая догадка
  if (/^oz|ozon/i.test(s))       return 'OZ';
  return null;
})(mpRaw);
if (!PLAT) throw new Error('Не распознана площадка в именованном диапазоне muff_mp. Значение: "' + mpRaw + '"');
log('platform detect', 'muff_mp="' + mpRaw + '" -> ' + PLAT);

// === Кабинет ТОЛЬКО из именованного muff_cabs
var cabinet = (REF && typeof REF.getCabinetControlValue === 'function') ? REF.getCabinetControlValue() : '';
if (!cabinet) throw new Error('Не выбран кабинет (именованный muff_cabs)');
log('cabinet detect', cabinet);


    // Текущий режим «Ключи»
    var mode = (function getMode() {
      var lastRow = shPar.getLastRow(), lastCol = shPar.getLastColumn();
      if (lastRow < 2 || lastCol < 11) return 'Артикулы';
      var rng = shPar.getRange(1, 11, lastRow, Math.min(2, lastCol - 10)).getDisplayValues(); // K:L
      var m = 'Артикулы';
      for (var i = 0; i < rng.length; i++) {
        var key = String(rng[i][0] || '').trim().toLowerCase();
        if (key === 'ключи') {
          var v = String(rng[i][1] || '').trim();
          m = (v === 'Названия') ? 'Названия' : 'Артикулы';
          break;
        }
      }
      return m;
    })();

    log('START', 'platform=' + PLAT + ', cabinet=' + cabinet + ', mode=' + mode);

    // РЕЗОЛВЕР: читаем лист артикулов площадки и строим map’ы по текущему кабинету
    var resolver = buildIdResolverByPlatformCabinet_(PLAT, cabinet);
    log('resolver built',
        'byDisplay=' + resolver.byDisplay.size + ', byVendor=' + resolver.byVendor.size +
        ', byOffer=' + resolver.byOffer.size + ', byNm=' + resolver.byNm.size);

    // Собираем G/H до последней непустой G
    var lastCalcRow = shCalc.getLastRow();
    var gVals = shCalc.getRange(4, 7, Math.max(lastCalcRow - 3, 1), 1).getDisplayValues(); // G4:G
    var hVals = shCalc.getRange(4, 8, Math.max(lastCalcRow - 3, 1), 1).getValues();        // H4:H

    var lastIdx = -1;
    for (var r = gVals.length - 1; r >= 0; r--) {
      if (String(gVals[r][0] || '').trim() !== '') { lastIdx = r; break; }
    }
    if (lastIdx < 0) { log('no rows', 'G пусто'); ss.toast('Нет строк для отправки (G пусто)', 'Готово', 4); return; }

    log('scan rows', 'rows= ' + (lastIdx + 1));

    var payloadOZ = [];
    var payloadWB = [];
    var stats = { resolved: 0, unresolved: 0, badPrice: 0, emptyKey: 0 };

    for (var i = 0; i <= lastIdx; i++) {
      var keyDisp = String(gVals[i][0] || '').trim();
      var priceRaw = hVals[i][0];
      if (!keyDisp) { stats.emptyKey++; continue; }

      var price = Number(priceRaw);
      if (!(isFinite(price) && price > 0)) { stats.badPrice++; continue; }

      if (PLAT === 'OZ') {
        var offerId = resolveOzonOfferId_(keyDisp, mode, resolver);
        if (offerId) {
          payloadOZ.push({ offer_id: offerId, price: String(Math.round(price)) });
          stats.resolved++;
        } else {
          stats.unresolved++;
        }
      } else {
        var nm = resolveWbNmId_(keyDisp, mode, resolver);
        if (nm) {
          payloadWB.push({ nmID: Number(nm), price: Math.round(Number(price)), discount: 0 });
          stats.resolved++;
        } else {
          stats.unresolved++;
        }
      }
    }

    log('collect done',
        'resolved=' + stats.resolved + ', unresolved=' + stats.unresolved +
        ', badPrice=' + stats.badPrice + ', emptyKey=' + stats.emptyKey);

    // ===== OZON =====
    if (PLAT === 'OZ') {
      if (!payloadOZ.length) { ss.toast('OZON: нет валидных строк для отправки', 'Готово', 4); return; }
      log('payload OZ (first 5)', JSON.stringify(payloadOZ.slice(0, 5)));
      var oz = new OZONAPI(cabinet);
      try {
        var resOZ = oz.setPrices(payloadOZ);
        log('OZ setPrices OK', 'sent=' + payloadOZ.length + ', result_len=' + (resOZ && resOZ.length || 0));
        ss.toast('OZON: отправлено цен: ' + payloadOZ.length, 'Готово', 6);
      } catch (e) {
        log('OZ setPrices FAIL', (e && e.message) ? e.message : String(e));
        throw new Error('OZON setPrices failed: ' + ((e && e.message) || e));
      }
      return;
    }

    // ===== WB =====
    if (!payloadWB.length) { ss.toast('WB: нет валидных строк для отправки', 'Готово', 4); return; }
    log('payload WB (first 5)', JSON.stringify(payloadWB.slice(0, 5)));

var tokenWB = (REF && REF.pickWBToken) ? REF.pickWBToken(cabinet) : null;
if (!tokenWB) throw new Error('WB: не найден ни один токен с доступом "Цены и скидки" для кабинета ' + cabinet);


    log('WB token picked', 'role=prices');

    try {
      var wbClient = new WB(tokenWB);
      var resWB = wbClient.setPrices(payloadWB, { batchSize: 500 });

      // Подхватим возможные uploadId(ы) из ответа
      var uploadIds = [];
      if (resWB) {
        if (Array.isArray(resWB.uploadIds)) uploadIds = resWB.uploadIds.slice();
        else if (resWB.data && typeof resWB.data.id !== 'undefined') uploadIds = [resWB.data.id];
        else if (resWB.id) uploadIds = [resWB.id];
      }

      log('WB setPrices OK', 'sent=' + payloadWB.length + ', uploadIds=' + JSON.stringify(uploadIds));
      ss.toast('WB: отправлено цен: ' + payloadWB.length + (uploadIds.length ? (' | uploadId=' + uploadIds.join(',')) : ''), 'Готово', 7);

      // 🔎 Отладка + выгрузка лога в «🛠 Тех. лог» (с H)
      WB_debugPriceUpload_(tokenWB, uploadIds, payloadWB, cabinet);

    } catch (e2) {
      log('WB setPrices FAIL', (e2 && e2.message) ? e2.message : String(e2));
      throw new Error('WB setPrices failed: ' + ((e2 && e2.message) || e2));
    }

  } catch (err) {
    console.log('[sendPricesFromCalculatorFast] ERROR: ' + ((err && err.stack) || (err && err.message) || String(err)));
    SpreadsheetApp.getActive().toast('Ошибка отправки цен: ' + (err && err.message || err), 'Ошибка', 8);
    throw err;
  }
}

/** ====== РЕЗОЛВЕРЫ (по листам артикулов) ====== */
function buildIdResolverByPlatformCabinet_(PLAT, cabinet) {
  var t0 = Date.now();
  var LOG_NS = 'buildIdResolverByPlatformCabinet_';
  function log(msg, extra) {
    var ms = String(Date.now() - t0).padStart(6, ' ');
    console.log('[' + ms + ' ms][' + LOG_NS + '] ' + msg + (extra ? ' | ' + extra : ''));
  }

  var ss = SpreadsheetApp.getActive();
  var artsSheet = ss.getSheetByName(PLAT === 'WB' ? REF.SHEETS.ARTS_WB : REF.SHEETS.ARTS_OZ);
  if (!artsSheet) { log('no sheet'); return { byDisplay: new Map(), byVendor: new Map(), byOffer: new Map(), byNm: new Map() }; }

  var lastRow = artsSheet.getLastRow();
  if (lastRow < 2) { log('no data rows'); return { byDisplay: new Map(), byVendor: new Map(), byOffer: new Map(), byNm: new Map() }; }

  var vals = artsSheet.getRange(2, 1, lastRow - 1, 12).getDisplayValues(); // A:L
  var byDisplay = new Map(); // строка показа «как в калькуляторе» -> {offer_id?, nmID?}
  var byVendor  = new Map(); // vendorCode (WB B) -> nmID
  var byOffer   = new Map(); // offer_id (OZ B) -> offer_id
  var byNm      = new Map(); // nmID (WB K) -> nmID

  var rowsSeen = 0;
  vals.forEach(function (row) {
    var cab = String(row[0] || '').trim();
    if (cab !== cabinet) return;
    rowsSeen++;

    var B = String(row[1] || '').trim();   // OZ: Артикул(offer_id) | WB: Артикул продавца(vendor)
    var K = String(row[10] || '').trim();  // OZ: SKU              | WB: nmID
    var L = String(row[11] || '').trim();  // Наименование

    if (PLAT === 'OZ') {
      var offer = B;
      if (offer) byOffer.set(offer, offer);

      var disp = ((K || offer) ? (K || offer) : '') + (L ? (' | ' + L) : '');
      disp = disp.trim();
      if (disp) byDisplay.set(disp, { offer_id: offer });
      if (offer) byDisplay.set(offer, { offer_id: offer });

    } else { // WB
      var vendor = B;
      var nmID   = K;
      if (vendor && nmID) byVendor.set(vendor, nmID);
      if (nmID) byNm.set(nmID, nmID);

      var dispWB = (nmID ? nmID : '') + (L ? (' | ' + L) : '');
      dispWB = dispWB.trim();
      if (dispWB && nmID) byDisplay.set(dispWB, { nmID: nmID });

      if (vendor && nmID) byDisplay.set(vendor, { nmID: nmID });
      if (nmID) byDisplay.set(nmID, { nmID: nmID });
    }
  });

  log('built',
      'rowsSeen=' + rowsSeen +
      ', byDisplay=' + byDisplay.size +
      ', byVendor=' + byVendor.size +
      ', byOffer='  + byOffer.size  +
      ', byNm='     + byNm.size);

  return { byDisplay: byDisplay, byVendor: byVendor, byOffer: byOffer, byNm: byNm };
}

function resolveOzonOfferId_(keyDisp, mode, R) {
  var off = String(keyDisp || '').trim();
  if (mode === 'Артикулы') {
    if (off) return off;
  }
  var rec = R.byDisplay.get(off);
  if (rec && rec.offer_id) return rec.offer_id;
  return '';
}

function resolveWbNmId_(keyDisp, mode, R) {
  var s = String(keyDisp || '').trim();
  if (mode === 'Названия') {
    var rec = R.byDisplay.get(s);
    if (rec && rec.nmID) return rec.nmID;

    var left = s.split('|')[0].trim();
    if (R.byNm.has(left)) return left;
    return '';
  } else {
    if (R.byVendor.has(s)) return R.byVendor.get(s);
    if (R.byNm.has(s))     return s;
    return '';
  }
}

/* =========================================================
 * ========= WB DEBUG + ВЫГРУЗКА ЛОГА В «🛠 Тех. лог» =======
 * ========================================================= */

/**
 * Собирает развернутый лог по загрузке цен WB и выгружает в лист «🛠 Тех. лог» с колонки H.
 * - Очищает область H:… по числу колонок таблицы (все строки листа), ставит заголовки в 1-й строке.
 * - Данные пишет со 2-й строки.
 * - На каждую позицию (nmID) одна строка с контекстом uploadID и статусом.
 *
 * @param {string} token       WB токен (роль «Цены и скидки, Аналитика»)
 * @param {number[]} uploadIds Идентификаторы загрузки (массив, может быть пустым)
 * @param {{nmID:number, price:number, discount:number}[]} payloadWB — то, что отправляли
 * @param {string=} cabinet    Опционально — имя кабинета (для наглядности в логе)
 */
function WB_debugPriceUpload_(token, uploadIds, payloadWB, cabinet) {
  var LOG_NS = 'WB_DEBUG';
  function clog(label, extra) {
    console.log('[' + LOG_NS + '] ' + label + (extra ? ' | ' + extra : ''));
  }

  if (!token) { clog('skip', 'no token'); return; }

  // === 0) Базовые коллекции ===
  var now = new Date();
  var plat = 'WB';
  var allNmIDs = [];
  try {
    var seen = new Set();
    for (var i = 0; i < payloadWB.length; i++) {
      var n = Number(payloadWB[i] && payloadWB[i].nmID);
      if (n && !seen.has(n)) { seen.add(n); allNmIDs.push(n); }
    }
  } catch (_) {}
  clog('nmIDs collected', 'count=' + allNmIDs.length);

  // Карта по nmID с будущими полями (заполняем постепенно)
  var recMap = {}; // nmID -> {ts, plat, cabinet, uploadID, status, nmID, error, quarantine, price, discount, discountedPrice, clubDiscountedPrice}
  function ensureRec(nm, upId, statusStr) {
    if (!recMap[nm]) recMap[nm] = {
      ts: now, plat: plat, cabinet: cabinet || '',
      uploadID: (upId || ''), status: (statusStr || ''),
      nmID: nm, error: '', quarantine: '',
      price: '', discount: '', discountedPrice: '', clubDiscountedPrice: ''
    };
    if (upId && !recMap[nm].uploadID) recMap[nm].uploadID = upId;
    if (statusStr && !recMap[nm].status) recMap[nm].status = statusStr;
    return recMap[nm];
  }

  // === 1) Если есть uploadIds — проверим статусы и вытащим детализацию ===
  var hasUploads = Array.isArray(uploadIds) && uploadIds.length > 0;
  if (hasUploads) {
    var statusName = function(n){ // для history/tasks
      switch(Number(n)){ case 3: return 'processed'; case 5: return 'partial'; case 6: return 'errors'; default: return String(n); }
    };
    for (var u = 0; u < uploadIds.length; u++) {
      var upId = Number(uploadIds[u]);
      if (!upId) continue;

      var statusStr = '';
      // До 10 попыток, интервал ~1.2s
      for (var attempt = 1; attempt <= 10; attempt++) {
        var history = WB_get_(token, 'https://discounts-prices-api.wildberries.ru/api/v2/history/tasks?uploadID=' + upId, 'history/tasks');
        if (history && history.data && typeof history.data.status !== 'undefined') {
          statusStr = statusName(history.data.status);
          clog('history state', 'uploadID=' + upId + ', status=' + statusStr);
          if (['processed','partial','errors'].indexOf(statusStr) >= 0) break;
        } else {
          clog('history state', 'no data for uploadID=' + upId);
        }

        var buffer = WB_get_(token, 'https://discounts-prices-api.wildberries.ru/api/v2/buffer/tasks?uploadID=' + upId, 'buffer/tasks');
        if (buffer && buffer.data && typeof buffer.data.status !== 'undefined') {
          var bst = Number(buffer.data.status); // 1=in progress
          if (bst === 1) { statusStr = 'in_progress'; clog('buffer state', 'uploadID=' + upId + ', status=1'); }
          else { clog('buffer state', 'uploadID=' + upId + ', status=' + bst); }
          if (bst !== 1) break;
        }

        Utilities.sleep(1200);
      }

      // Детали (history первичен; если нет — buffer)
      var details = WB_get_(token,
        'https://discounts-prices-api.wildberries.ru/api/v2/history/goods/task?uploadID=' + upId + '&limit=1000&offset=0',
        'history/goods/task');
      if (!(details && details.data && Array.isArray(details.data.listGoods))) {
        details = WB_get_(token,
          'https://discounts-prices-api.wildberries.ru/api/v2/buffer/goods/task?uploadID=' + upId + '&limit=1000&offset=0',
          'buffer/goods/task');
      }

      if (details && details.data && Array.isArray(details.data.listGoods)) {
        details.data.listGoods.forEach(function(g){
          var nm = Number(g && g.nmID);
          if (!nm) return;
          var err = String((g.errorText || g.error || '')).trim();
          var R = ensureRec(nm, upId, statusStr || '');
          if (err) R.error = err;
        });
      }

      // Если в деталях ничего не было — всё равно положим строки по исходному списку nmID
      if ((!details || !details.data || !Array.isArray(details.data.listGoods)) && allNmIDs.length) {
        allNmIDs.forEach(function(nm){
          ensureRec(nm, upId, statusStr || '');
        });
      }
    }
  } else {
    // Без uploadID: ведём строки по исходному набору nmID
    allNmIDs.forEach(function(nm){ ensureRec(nm, '', 'no_upload'); });
  }

  // Если вообще ничего не собралось — защитно выведем хотя бы «пустышки»
  if (!Object.keys(recMap).length && allNmIDs.length) {
    allNmIDs.forEach(function(nm){ ensureRec(nm, '', 'no_upload'); });
  }

  // === 2) Карантин и текущие цены ===
  var nmList = Object.keys(recMap).map(function(k){ return Number(k); });
  if (nmList.length) {
    var quarantineMap = WB_fetchQuarantine_(token, nmList);
    var pricesMap     = WB_fetchPrices_(token, nmList);
    nmList.forEach(function(nm){
      var R = recMap[nm];
      if (!R) return;
      if (quarantineMap[nm]) R.quarantine = quarantineMap[nm];
      if (pricesMap[nm]) {
        R.price = pricesMap[nm].price;
        R.discount = pricesMap[nm].discount;
        if (typeof pricesMap[nm].discountedPrice !== 'undefined') R.discountedPrice = pricesMap[nm].discountedPrice;
        if (typeof pricesMap[nm].clubDiscountedPrice !== 'undefined') R.clubDiscountedPrice = pricesMap[nm].clubDiscountedPrice;
      }
    });
  }

  // === 3) Выгрузка в «🛠 Тех. лог» с H ===
  var HEAD = [
    'Время', 'Площадка', 'Кабинет', 'UploadID', 'Статус',
    'nmID', 'Ошибка', 'Карантин',
    'Текущая цена', 'Скидка', 'Цена со скидкой', 'Клубная цена'
  ];

  var rows = Object.keys(recMap).sort(function(a,b){ return Number(a)-Number(b); }).map(function(k){
    var r = recMap[k];
    return [
      r.ts, r.plat, r.cabinet, r.uploadID, r.status,
      r.nmID, r.error, r.quarantine,
      r.price, r.discount, r.discountedPrice, r.clubDiscountedPrice
    ];
  });

  WB_writeTechLogWB_(HEAD, rows);
  clog('written to sheet', 'rows=' + rows.length);
}

/** ===== Низкоуровневые вызовы WB (не логируем токен!) ===== */
function WB_post_(token, url, bodyObj, tag) {
  var resp, code, txt;
  try {
    resp = UrlFetchApp.fetch(url, {
      method: 'post',
      payload: JSON.stringify(bodyObj || {}),
      contentType: 'application/json',
      headers: { 'Authorization': token },
      muteHttpExceptions: true
    });
    code = resp.getResponseCode();
    txt  = resp.getContentText() || '';
  } catch (e) {
    console.log('[WB_POST][' + (tag||'') + '] ' + url + ' | EXC: ' + (e && e.message || e));
    return null;
  }
  var trimmed = txt.length > 2000 ? (txt.slice(0, 2000) + '…') : txt;
  console.log('[WB_POST][' + (tag||'') + '] ' + url + ' | code=' + code + ' | body=' + trimmed);
  if (code >= 200 && code < 300) {
    try { return JSON.parse(txt); } catch (_) { return { raw: txt }; }
  }
  return null;
}

function WB_get_(token, url, tag) {
  var resp, code, txt;
  try {
    resp = UrlFetchApp.fetch(url, {
      method: 'get',
      contentType: 'application/json',
      headers: { 'Authorization': token },
      muteHttpExceptions: true
    });
    code = resp.getResponseCode();
    txt  = resp.getContentText() || '';
  } catch (e) {
    console.log('[WB_GET][' + (tag||'') + '] ' + url + ' | EXC: ' + (e && e.message || e));
    return null;
  }
  var trimmed = txt.length > 2000 ? (txt.slice(0, 2000) + '…') : txt;
  console.log('[WB_GET][' + (tag||'') + '] ' + url + ' | code=' + code + ' | body=' + trimmed);
  if (code >= 200 && code < 300) {
    try { return JSON.parse(txt); } catch (_) { return { raw: txt }; }
  }
  return null;
}

/** Возвращает Map nmID->reason для карантина цен */
function WB_fetchQuarantine_(token, nmIDs) {
  var out = {};
  if (!nmIDs || !nmIDs.length) return out;

  for (var i = 0; i < nmIDs.length; i += 1000) {
    var chunk = nmIDs.slice(i, i + 1000);
    var res = WB_post_(token,
      'https://discounts-prices-api.wildberries.ru/api/v2/quarantine/goods',
      { nmIDs: chunk }, 'quarantine/goods');
    if (res && res.data && Array.isArray(res.data.listGoods)) {
      res.data.listGoods.forEach(function(g){
        var nm = Number(g && g.nmID);
        if (!nm) return;
        var reason = String(g.reason || g.error || '').trim();
        if (reason) out[nm] = reason;
      });
    }
  }
  return out;
}

/** Возвращает Map nmID->{price, discount, discountedPrice?, clubDiscountedPrice?} */
function WB_fetchPrices_(token, nmIDs) {
  var out = {};
  if (!nmIDs || !nmIDs.length) return out;

  for (var i = 0; i < nmIDs.length; i += 1000) {
    var chunk = nmIDs.slice(i, i + 1000);
    var res = WB_post_(token,
      'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter',
      { nmIDs: chunk }, 'list/goods/filter');
    if (res && res.data && Array.isArray(res.data.listGoods)) {
      res.data.listGoods.forEach(function(g){
        var nm = Number(g && g.nmID);
        if (!nm) return;
        out[nm] = {
          price: g.price,
          discount: g.discount,
          discountedPrice: g.discountedPrice,
          clubDiscountedPrice: g.clubDiscountedPrice
        };
      });
    }
  }
  return out;
}

/**
 * Запись таблицы лога в лист «🛠 Тех. лог» с колонки H.
 * Перед записью: очищаем область H:… (все строки листа, по ширине заголовков),
 * затем пишем заголовки в 1-й строке и данные со 2-й.
 */
function WB_writeTechLogWB_(HEAD, rows) {
  var ss = SpreadsheetApp.getActive();
  var name = '🛠 Тех. лог';
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);

  var startCol = 8; // H
  var width = HEAD.length;

  // Очистка области вставки по всей высоте листа (только контент)
  var maxRows = sh.getMaxRows();
  sh.getRange(1, startCol, maxRows, width).clearContent();

  // Заголовки (1-я строка)
  sh.getRange(1, startCol, 1, width).setValues([HEAD]).setFontWeight('bold');

  // Данные (со 2-й строки)
  if (rows && rows.length) {
    sh.getRange(2, startCol, rows.length, width).setValues(rows);
  }

  // Формат времени
  sh.getRange(2, startCol, Math.max(1, rows.length), 1).setNumberFormat('dd.mm.yyyy HH:mm:ss');

  // Авто-ширина для зоны лога
  for (var c = 0; c < width; c++) {
    sh.autoResizeColumn(startCol + c);
  }
}
