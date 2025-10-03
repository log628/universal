/** Универсальная отправка цен из «⚖️ Калькулятор» в включённую площадку (⚙️ Параметры!I2) */
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
    var platRaw = String(shPar.getRange('I2').getDisplayValue() || '').trim().toUpperCase();
    var PLAT = (platRaw === 'WILDBERRIES' || platRaw === 'WB') ? 'WB' : 'OZ';

    // Текущий кабинет
    var cabinet = String(shCalc.getRange('B3:E4').getDisplayValue() || '').trim();
    if (!cabinet) throw new Error('Не выбран кабинет (⚖️ Калькулятор!B3:E4)');

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

    // ===== WB ветка: берём токен роли "Цены и скидки, Аналитика" и шлём через класс WB =====
    if (!payloadWB.length) { ss.toast('WB: нет валидных строк для отправки', 'Готово', 4); return; }
    log('payload WB (first 5)', JSON.stringify(payloadWB.slice(0, 5)));

    var tokenWB = (REF && REF.pickWBToken) ? REF.pickWBToken(cabinet, 'prices', true) : null;
    if (!tokenWB) throw new Error('WB: не найден токен с ролью "Цены и скидки, Аналитика" для кабинета ' + cabinet);

    // НЕ логируем токен в консоль!
    log('WB token picked', 'role=prices');

    try {
      var wbClient = new WB(tokenWB);
      var resWB = wbClient.setPrices(payloadWB, { batchSize: 500 });
      log('WB setPrices OK', 'sent=' + (resWB && resWB.count || payloadWB.length) + ', tasks=' + ((resWB && resWB.uploadIds && resWB.uploadIds.length) || 0));
      ss.toast('WB: отправлено цен: ' + (resWB && resWB.count || payloadWB.length) +
               (resWB && resWB.uploadIds && resWB.uploadIds.length ? (', tasks=' + resWB.uploadIds.join(',')) : ''),
               'Готово', 7);
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

      // строка показа для режима «Названия»: (SKU || Артикул) + ' | ' + Наименование
      var disp = ((K || offer) ? (K || offer) : '') + (L ? (' | ' + L) : '');
      disp = disp.trim();
      if (disp) byDisplay.set(disp, { offer_id: offer });
      if (offer) byDisplay.set(offer, { offer_id: offer }); // дубль на «чистый» артикул

    } else { // WB
      var vendor = B;
      var nmID   = K;
      if (vendor && nmID) byVendor.set(vendor, nmID);
      if (nmID) byNm.set(nmID, nmID);

      // строка показа для режима «Названия»: nmID + ' | ' + Наименование
      var dispWB = (nmID ? nmID : '') + (L ? (' | ' + L) : '');
      dispWB = dispWB.trim();
      if (dispWB && nmID) byDisplay.set(dispWB, { nmID: nmID });

      // и «чистые» ключи
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
  // Если режим «Артикулы» — принимаем G как offer_id; иначе пробуем карту byDisplay
  var off = String(keyDisp || '').trim();
  if (mode === 'Артикулы') {
    // На OZ offer_id = «Артикул» (колонка B выгрузки). Доверяем «Калькулятору».
    if (off) return off;
  }
  var rec = R.byDisplay.get(off);
  if (rec && rec.offer_id) return rec.offer_id;
  return '';
}

function resolveWbNmId_(keyDisp, mode, R) {
  var s = String(keyDisp || '').trim();
  if (mode === 'Названия') {
    // G = «nmID | Наименование» — пробуем карту, затем парсим левую часть до «|»
    var rec = R.byDisplay.get(s);
    if (rec && rec.nmID) return rec.nmID;

    var left = s.split('|')[0].trim();
    if (R.byNm.has(left)) return left;
    return '';
  } else {
    // «Артикулы» (для WB это «Артикул продавца») — ищем в byVendor; иногда G уже nmID
    if (R.byVendor.has(s)) return R.byVendor.get(s);
    if (R.byNm.has(s))     return s;
    return '';
  }
}
