/** =========================================================
 * Универсальная отправка цен из «⚖️ Калькулятор» в включённую площадку (⚙️ Параметры!I2)
 * + Полный WB-лог со статусами/ошибками/карантином/OK? в лист «🛠 Тех. лог» (с H)
 * + Учитывает size-pricing на WB (editableSizePrice: true → upload/task/size)
 * + Пишет в лог Артикул продавца (vendorCode) и блокировки Sale/остатки
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

    // === Площадка ТОЛЬКО из именованного muff_mp (REF.NAMED.MP_CTRL)
    var rMP   = (REF && REF.NAMED && REF.NAMED.MP_CTRL) ? ss.getRangeByName(REF.NAMED.MP_CTRL) : null;
    var mpRaw = rMP ? String(rMP.getDisplayValue() || '').trim() : '';
    var PLAT  = (function (s) {
      if (REF && typeof REF.platformCanon === 'function') {
        var c = REF.platformCanon(s);            // 'OZ' | 'WB' | null
        if (c) return c;
      }
      if (/wb|wildberries/i.test(s)) return 'WB';
      if (/^oz|ozon/i.test(s))       return 'OZ';
      return null;
    })(mpRaw);
    if (!PLAT) throw new Error('Не распознана площадка в именованном диапазоне muff_mp. Значение: "' + mpRaw + '"');
    log('platform detect', 'muff_mp="' + mpRaw + '" -> ' + PLAT);

    // === Кабинет ТОЛЬКО из именованного muff_cabs
    var cabinet = (REF && typeof REF.getCabinetControlValue === 'function') ? REF.getCabinetControlValue() : '';
    if (!cabinet) throw new Error('Не выбран кабинет (именованный muff_cabs)');
    log('cabinet detect', cabinet);

    log('START', 'platform=' + PLAT + ', cabinet=' + cabinet + ', mode=Артикулы');

    // === РЕЗОЛВЕР: лист артикулов площадки → словари по кабинету (WB: vendorCode→nmID; OZ: offer_id passthrough)
    var resolver = buildIdResolverByPlatformCabinet_(PLAT, cabinet);
    log('resolver built',
        'byVendor=' + resolver.byVendor.size +
        ', byNm=' + resolver.byNm.size + ', byOffer=' + resolver.byOffer.size);

    // === Собираем G/H до последней непустой G
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
      var key = String(gVals[i][0] || '').trim();    // G: «Артикулы»
      var priceRaw = hVals[i][0];
      if (!key) { stats.emptyKey++; continue; }

      var price = Number(priceRaw);
      if (!(isFinite(price) && price > 0)) { stats.badPrice++; continue; }

      if (PLAT === 'OZ') {
        // G = offer_id
        var offerId = key;
        if (offerId) {
          payloadOZ.push({ offer_id: offerId, price: String(Math.round(price)) });
          stats.resolved++;
        } else {
          stats.unresolved++;
        }
      } else {
        // WB: G = vendorCode; резолвим nmID, либо берём G как nmID если так и введено
        var nm = '';
        if (resolver.byVendor.has(key)) {
          nm = resolver.byVendor.get(key);       // vendorCode -> nmID
        } else if (resolver.byNm.has(key)) {
          nm = key;                               // допустим G уже nmID
        }
        if (nm) {
          payloadWB.push({ nmID: Number(nm), price: Math.round(Number(price)) });
          stats.resolved++;
        } else {
          stats.unresolved++;
        }
      }
    }

    log('collect done', 'resolved=' + stats.resolved + ', unresolved=' + stats.unresolved +
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

    // 1) Получим актуальную информацию по товарам → решаем, где size-pricing
    var nmList = payloadWB.map(function(x){ return Number(x.nmID); });
    var infoMap = WB_fetchProductsInfo_(tokenWB, nmList); // nmID -> {vendorCode, sizes[], editableSizePrice, price, discount, isBadTurnover}

    var sizeData = [];      // [{nmID, sizeID, price}]
    var productData = [];   // [{nmID, price}]
    var nmMode = {};        // nmID -> 'size' | 'product'

    nmList.forEach(function(nm){
      var inf = infoMap[nm];
      if (inf && inf.editableSizePrice && Array.isArray(inf.sizes) && inf.sizes.length) {
        var p = payloadWB.find(function(x){ return Number(x.nmID) === nm; }).price;
        inf.sizes.forEach(function(sz){ if (sz && typeof sz.sizeID !== 'undefined') sizeData.push({ nmID: nm, sizeID: Number(sz.sizeID), price: p }); });
        nmMode[nm] = 'size';
      } else {
        var p2 = payloadWB.find(function(x){ return Number(x.nmID) === nm; }).price;
        productData.push({ nmID: nm, price: p2 });
        nmMode[nm] = 'product';
      }
    });

    // 2) Отправка: товары и размеры отдельными задачами (чанк по лимитам)
    var uploadIds = [];
    if (productData.length) {
      var chunksP = chunk_(productData, 1000);
      for (var cp = 0; cp < chunksP.length; cp++) {
        var rP = WB_post_(tokenWB, 'https://discounts-prices-api.wildberries.ru/api/v2/upload/task', { data: chunksP[cp] }, 'upload/task');
        var idP = (rP && rP.data && typeof rP.data.id !== 'undefined') ? Number(rP.data.id) : null;
        if (idP) uploadIds.push(idP);
        Utilities.sleep(650);
      }
    }
    if (sizeData.length) {
      var chunksS = chunk_(sizeData, 1000);
      for (var cs = 0; cs < chunksS.length; cs++) {
        var rS = WB_post_(tokenWB, 'https://discounts-prices-api.wildberries.ru/api/v2/upload/task/size', { data: chunksS[cs] }, 'upload/task/size');
        var idS = (rS && rS.data && typeof rS.data.id !== 'undefined') ? Number(rS.data.id) : null;
        if (idS) uploadIds.push(idS);
        Utilities.sleep(650);
      }
    }

    log('WB uploads', 'sent products=' + productData.length + ', sizes=' + sizeData.length + ', uploadIds=' + JSON.stringify(uploadIds));
    ss.toast('WB: отправлено цен: ' + payloadWB.length + (uploadIds.length ? (' | uploadId=' + uploadIds.join(',')) : ''), 'Готово', 7);

    // 3) Подробный лог в «🛠 Тех. лог» (size-pricing + vendorCode + OK? + блокировки)
    WB_debugPriceUpload_(tokenWB, uploadIds, payloadWB, cabinet, { nmMode: nmMode });

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

  function norm(s){ return String(s||'').replace(/\[[^\]]*\]/g,'').trim().toLowerCase(); }
  function findIdx(hdr, variants){
    var set = (variants||[]).map(norm);
    for (var i=0;i<hdr.length;i++){
      var h = norm(hdr[i]);
      if (set.indexOf(h) !== -1) return i+1; // 1-based
    }
    return 0;
  }

  var ss = SpreadsheetApp.getActive();
  var artsSheet = ss.getSheetByName(PLAT === 'WB' ? REF.SHEETS.ARTS_WB : REF.SHEETS.ARTS_OZ);
  if (!artsSheet) {
    log('no sheet');
    return { byVendor:new Map(), byNm:new Map(), byOffer:new Map() };
  }

  var lastRow = artsSheet.getLastRow();
  var lastCol = artsSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) {
    log('no data rows');
    return { byVendor:new Map(), byNm:new Map(), byOffer:new Map() };
  }

  // читаем всю строку заголовков как есть
  var hdr = artsSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

  // базовые индексы
  var cCab  = findIdx(hdr, ['кабинет']) || 1; // A по умолчанию

  // OZ (только «Артикул» / offer_id)
  var cOfferOZ = findIdx(hdr, ['offer_id','артикул продавца','артикул']) || 2; // B fallback

  // WB (vendorCode и nmID (в твоей схеме nmID в колонке "SKU"))
  var cVendorWB = findIdx(hdr, ['артикул продавца','vendorcode','vendor_code']) || 2; // B
  var cNmWB     = findIdx(hdr, ['nmid','nm id','nm_id','nm','sku']) || 11;            // K/«SKU»

  var needCols = Math.max(cCab, cOfferOZ, cVendorWB, cNmWB);
  needCols = Math.min(needCols || lastCol, lastCol);
  var vals = artsSheet.getRange(2, 1, lastRow - 1, needCols).getDisplayValues();

  var byVendor  = new Map(); // WB: vendorCode -> nmID
  var byNm      = new Map(); // WB: nmID       -> nmID
  var byOffer   = new Map(); // OZ: offer_id   -> offer_id (на всякий случай)

  var rowsSeen = 0;

  for (var i=0;i<vals.length;i++){
    var row = vals[i];
    var cab = String(row[cCab-1] || '').trim();
    if (cab !== cabinet) continue;
    rowsSeen++;

    if (PLAT === 'OZ') {
      var offer = (cOfferOZ ? String(row[cOfferOZ-1] || '').trim() : '');
      if (offer) byOffer.set(offer, offer);
    } else {
      var vendor = (cVendorWB ? String(row[cVendorWB-1] || '').trim() : '');
      var nmID   = (cNmWB     ? String(row[cNmWB-1]     || '').trim() : '');
      if (vendor && nmID) byVendor.set(vendor, nmID);
      if (nmID)           byNm.set(nmID, nmID);
    }
  }

  log('built','rowsSeen=' + rowsSeen + ', byVendor=' + byVendor.size + ', byNm=' + byNm.size + ', byOffer=' + byOffer.size);

  return { byVendor: byVendor, byNm: byNm, byOffer: byOffer };
}

/* =========================================================
 * ======= WB DEBUG + ВЫГРУЗКА ЛОГА В «🛠 Тех. лог» (H) ======
 * ========================================================= */

function WB_debugPriceUpload_(token, uploadIds, payloadWB, cabinet, opts) {
  var LOG_NS = 'WB_DEBUG';
  function clog(label, extra) { console.log('[' + LOG_NS + '] ' + label + (extra ? ' | ' + extra : '')); }
  opts = opts || {};

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

  // nm→отправленная цена + режим (если есть)
  var sentPrice = {}; // nm -> price
  var nmMode    = opts.nmMode || {}; // nm -> 'size' | 'product'
  payloadWB.forEach(function(p){ var n = Number(p.nmID); if (n) sentPrice[n] = Number(p.price); });

  // Строки лога по nmID
  var recMap = {}; // nmID -> {...}
  function ensureRec(nm, upId, statusStr) {
    if (!recMap[nm]) recMap[nm] = {
      ts: now, plat: plat, cabinet: cabinet || '',
      uploadID: (upId || ''), status: (statusStr || ''),
      nmID: nm, vendor: '', error: '', quarantine: '',
      price: '', discount: '', discountedPrice: '', clubDiscountedPrice: '',
      sentPrice: (typeof sentPrice[nm] !== 'undefined') ? sentPrice[nm] : '',
      delta: '', equal: '', mode: (nmMode[nm] || ''),
      badTurnover: '', block: '',
      result: '', reason: ''
    };
    if (upId && !recMap[nm].uploadID) recMap[nm].uploadID = upId;
    if (statusStr && !recMap[nm].status) recMap[nm].status = statusStr;
    return recMap[nm];
  }

  // Предзаполнение
  var hasUploads = Array.isArray(uploadIds) && uploadIds.length > 0;
  var preStatus = hasUploads ? 'sent' : 'no_upload (runtime)';
  var preUploadId = hasUploads ? String(uploadIds[0] || '') : '';
  allNmIDs.forEach(function(nm){ ensureRec(nm, preUploadId, preStatus); });

  // === 1) uploadIds → статусы и ошибки ===
  if (hasUploads) {
    var statusName = function(n){ switch(Number(n)){ case 3: return 'processed'; case 5: return 'partial'; case 6: return 'errors'; case 4: return 'canceled'; default: return String(n); } };
    for (var u = 0; u < uploadIds.length; u++) {
      var upId = Number(uploadIds[u]);
      if (!upId) continue;

      var statusStr = '';
      for (var attempt = 1; attempt <= 10; attempt++) {
        var history = WB_get_(token, 'https://discounts-prices-api.wildberries.ru/api/v2/history/tasks?uploadID=' + upId, 'history/tasks');
        if (history && history.data && typeof history.data.status !== 'undefined') {
          statusStr = statusName(history.data.status);
          clog('history state', 'uploadID=' + upId + ', status=' + statusStr);
          if (['processed','partial','errors','canceled'].indexOf(statusStr) >= 0) break;
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
      var goodsArr = (details && details.data && (details.data.historyGoods || details.data.listGoods)) || [];
      if (!goodsArr.length) {
        var details2 = WB_get_(token,
          'https://discounts-prices-api.wildberries.ru/api/v2/buffer/goods/task?uploadID=' + upId + '&limit=1000&offset=0',
          'buffer/goods/task');
        goodsArr = (details2 && details2.data && (details2.data.historyGoods || details2.data.bufferGoods || details2.data.listGoods)) || [];
      }
      if (Array.isArray(goodsArr)) {
        goodsArr.forEach(function(g){
          var nm = Number(g && g.nmID); if (!nm) return;
          var R = ensureRec(nm, upId, statusStr || '');
          var err = String((g.errorText || g.error || '')).trim();
          if (err) R.error = err;
          if (/Sale due to high inventory|You can't change the item price/i.test(err)) R.block = 'sale_high_inventory';
        });
      }

      // Обновим общий статус
      allNmIDs.forEach(function(nm){ var R = ensureRec(nm, upId, statusStr || ''); });
    }
  }

  // === 2) Карантин и текущие цены + vendor + isBadTurnover ===
  var nmList = Object.keys(recMap).map(function(k){ return Number(k); });
  if (nmList.length) {
    var quarantineMap = WB_fetchQuarantine_(token, nmList);
    var pricesMap     = WB_fetchPrices_(token, nmList);

    nmList.forEach(function(nm){
      var R = recMap[nm]; if (!R) return;
      if (quarantineMap[nm]) R.quarantine = quarantineMap[nm];
      var P = pricesMap[nm];
      if (P) {
        R.vendor = P.vendorCode || R.vendor || '';
        if (!R.mode) R.mode = (P.editableSizePrice ? 'size' : 'product');
        R.price  = (typeof P.price !== 'undefined') ? P.price : R.price;
        R.discount = (typeof P.discount !== 'undefined') ? P.discount : R.discount;
        if (typeof P.discountedPrice !== 'undefined') R.discountedPrice = P.discountedPrice;
        if (typeof P.clubDiscountedPrice !== 'undefined') R.clubDiscountedPrice = P.clubDiscountedPrice;
        if (typeof P.isBadTurnover !== 'undefined') R.badTurnover = P.isBadTurnover ? 'true' : '';
      }

      // Δ и OK? с допуском ±1 ₽
      if (R.sentPrice !== '' && R.sentPrice != null && R.price !== '' && R.price != null) {
        var sp = Number(R.sentPrice), cp = Number(R.price);
        if (isFinite(sp) && isFinite(cp)) {
          var d = cp - sp; var EPS = 1;
          R.delta = Math.round(d * 100) / 100;
          R.equal = (!R.block && Math.abs(d) <= EPS) ? 'OK' : '';
        }
      }

      // Итоговый статус/причина
      if (R.error) { R.result = 'error'; R.reason = R.error; }
      else if (R.block) { R.result = 'blocked'; R.reason = R.block; }
      else if (R.quarantine) { R.result = 'blocked'; R.reason = 'quarantine'; }
      else if (R.equal === 'OK') {
        if (R.status === 'no_upload (runtime)' || R.status === 'sent') { R.result = 'unchanged'; R.reason = 'same_price'; }
        else { R.result = 'applied'; R.reason = ''; }
      } else {
        R.result = 'pending'; R.reason = '';
      }
    });
  }

  // === 3) Выгрузка в лист «🛠 Тех. лог» с H ===
  var HEAD = [
    'Время', 'Площадка', 'Кабинет', 'UploadID', 'Статус',
    'nmID', 'Артикул продавца', 'Ошибка', 'Карантин', 'Блок (Sale/остатки)', 'Плохая оборачиваемость',
    'Текущая цена', 'Скидка', 'Цена со скидкой', 'Клубная цена',
    'Отправленная цена', 'Δ', 'OK?', 'Режим', 'Результат', 'Причина'
  ];

  var rows = Object.keys(recMap)
    .sort(function(a,b){ return Number(a)-Number(b); })
    .map(function(k){
      var r = recMap[k];
      return [
        r.ts, r.plat, r.cabinet, r.uploadID, r.status,
        r.nmID, r.vendor, r.error, r.quarantine, r.block, r.badTurnover,
        r.price, r.discount, r.discountedPrice, r.clubDiscountedPrice,
        r.sentPrice, r.delta, r.equal, r.mode, r.result, r.reason
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

/** Возвращает краткое описание карантина по nmID (only our nmIDs) */
function WB_fetchQuarantine_(token, nmIDs) {
  var out = {};
  if (!nmIDs || !nmIDs.length) return out;

  var need = new Set(nmIDs.map(Number).filter(Boolean));
  var limit = 1000, offset = 0, guard = 0;

  while (need.size && guard < 100) {
    var url = 'https://discounts-prices-api.wildberries.ru/api/v2/quarantine/goods' +
              '?limit=' + limit + '&offset=' + offset;
    var res = WB_get_(token, url, 'quarantine/goods');
    if (!(res && res.data)) break;
    var list = res.data.quarantineGoods || [];
    if (!list.length) break;

    for (var i = 0; i < list.length; i++) {
      var g  = list[i];
      var nm = Number(g && g.nmID);
      if (!need.has(nm)) continue;

      var np = Number(g.newPrice), op = Number(g.oldPrice);
      var nd = Number(g.newDiscount), od = Number(g.oldDiscount);
      var diff = (typeof g.priceDiff !== 'undefined') ? g.priceDiff : (isFinite(np) && isFinite(op) ? (np - op) : '');
      out[nm] = (isFinite(op) && isFinite(np))
        ? ('old ' + op + ' → new ' + np + (isFinite(diff) ? ' (Δ ' + diff + ')' : '') + (isFinite(nd) ? ', disc ' + nd + '%' : ''))
        : 'in quarantine';
      need.delete(nm);
    }

    offset += list.length;
    guard++;
    Utilities.sleep(650); // лимиты WB
  }
  return out;
}

/** Возвращает Map nmID → {vendorCode, price, discount, discountedPrice?, clubDiscountedPrice?, editableSizePrice?, isBadTurnover?} */
function WB_fetchPrices_(token, nmIDs) {
  var out = {};
  if (!nmIDs || !nmIDs.length) return out;

  for (var i = 0; i < nmIDs.length; i += 1000) {
    var chunk = nmIDs.slice(i, i + 1000).map(Number).filter(Boolean);
    var res = WB_post_(token, 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter', { nmList: chunk }, 'list/goods/filter');
    if (!(res && res.data && Array.isArray(res.data.listGoods))) continue;

    res.data.listGoods.forEach(function(g){
      var nm = Number(g && g.nmID);
      if (!nm) return;
      var priceTop = (typeof g.price !== 'undefined') ? g.price : null;
      var size0 = (Array.isArray(g.sizes) && g.sizes.length) ? g.sizes[0] : null;
      var price = (priceTop != null) ? priceTop : (size0 && typeof size0.price !== 'undefined' ? size0.price : '');
      var discountedPrice = (size0 && typeof size0.discountedPrice !== 'undefined') ? size0.discountedPrice : (price != null && typeof g.discount === 'number' ? Math.round(price * (100 - g.discount) / 100) : '');
      var clubDiscountedPrice = (size0 && typeof size0.clubDiscountedPrice !== 'undefined') ? size0.clubDiscountedPrice : '';
      out[nm] = {
        vendorCode: g.vendorCode,
        price: price,
        discount: (typeof g.discount === 'number') ? g.discount : '',
        discountedPrice: discountedPrice,
        clubDiscountedPrice: clubDiscountedPrice,
        editableSizePrice: !!g.editableSizePrice,
        isBadTurnover: !!g.isBadTurnover
      };
    });
    Utilities.sleep(650);
  }
  return out;
}

/** Получить полную инфу о товарах (для выбора режима size/product): vendorCode, sizes[], editableSizePrice, isBadTurnover */
function WB_fetchProductsInfo_(token, nmIDs) {
  var out = {};
  if (!nmIDs || !nmIDs.length) return out;

  for (var i = 0; i < nmIDs.length; i += 1000) {
    var chunk = nmIDs.slice(i, i + 1000).map(Number).filter(Boolean);
    var res = WB_post_(token, 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter', { nmList: chunk }, 'list/goods/filter(full)');
    if (!(res && res.data && Array.isArray(res.data.listGoods))) continue;

    res.data.listGoods.forEach(function(g){
      var nm = Number(g && g.nmID);
      if (!nm) return;
      out[nm] = {
        vendorCode: g.vendorCode,
        editableSizePrice: !!g.editableSizePrice,
        sizes: Array.isArray(g.sizes) ? g.sizes.slice() : [],
        price: (typeof g.price !== 'undefined') ? g.price : undefined,
        discount: (typeof g.discount === 'number') ? g.discount : undefined,
        isBadTurnover: !!g.isBadTurnover
      };
    });
    Utilities.sleep(650);
  }
  return out;
}

/** Запись таблицы лога в лист «🛠 Тех. лог» с колонки H */
function WB_writeTechLogWB_(HEAD, rows) {
  var ss = SpreadsheetApp.getActive();
  var name = '🛠 Тех. лог';
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);

  var startCol = 8; // H
  var width = HEAD.length;

  // Очистка области по всей высоте листа (только контент)
  var maxRows = sh.getMaxRows();
  sh.getRange(1, startCol, maxRows, width).clearContent();

  // Заголовки
  sh.getRange(1, startCol, 1, width).setValues([HEAD]).setFontWeight('bold');

  // Данные
  if (rows && rows.length) {
    sh.getRange(2, startCol, rows.length, width).setValues(rows);
  }

  // Формат времени
  sh.getRange(2, startCol, Math.max(1, (rows && rows.length) || 1), 1).setNumberFormat('dd.mm.yyyy HH:mm:ss');

  // Авто-ширина
  for (var c = 0; c < width; c++) sh.autoResizeColumn(startCol + c);
}

/* ========================= УТИЛИТЫ ========================= */
function chunk_(arr, n) {
  var out = []; if (!arr || !arr.length) return out; n = Math.max(1, n|0);
  for (var i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n));
  return out;
}
