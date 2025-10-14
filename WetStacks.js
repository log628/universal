/** =========================================================
 *  МЕГА ОСТАТКИ (OZON) — v1/analytics/stocks, без фильтров
 *  - Берём выбранные кабинеты из ⚙️ Параметры (D="OZON", H=TRUE)
 *  - По каждому кабинету: делаем отчёт по артикулам → offers[]
 *  - Маппим offer → sku через getArtsDetails (и фолбэк getSkusByOffers)
 *  - Вызываем /v1/analytics/stocks батчами по 100 SKU (как в OZONAPI)
 *  - НИЧЕГО не фильтруем в коде; пишем все поля, что вернул метод
 *  - Динамический заголовок = объединение всех ключей (плюс cabinet, offer_id, sku)
 *  - Пишем на лист "Мега Остатки" с 1-й строки заголовок, со 2-й — данные
 * ========================================================= */

function getMegaStocks_OZ_Analytics_NoFilters() {
  const ss = SpreadsheetApp.getActive();
  const SHEET_NAME = 'Мега Остатки';
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  const cabinets = _mega_getSelectedOZCabinets_(); // массив имён кабинетов
  if (!cabinets.length) {
    ss.toast('Нет отмеченных OZON-кабинетов в ⚙️ Параметры (колонка H)', 'Мега Остатки', 5);
    _mega_clearAll_(sh);
    return;
  }

  // Аккумуляторы “сырого” ответа
  const rows = [];               // массив плоских объектов
  const headerOrder = [];        // порядок появления ключей
  const headerSet = new Set();

  // Вспомогалки
  const pushFlat = (obj) => {
    const flat = _mega_flat_(obj);
    for (const k of Object.keys(flat)) {
      if (!headerSet.has(k)) { headerSet.add(k); headerOrder.push(k); }
    }
    rows.push(flat);
  };

  // Основной обход по кабинетам
  for (let i = 0; i < cabinets.length; i++) {
    const cab = cabinets[i];
    ss.toast('Остатки analytics/stocks: ' + cab, 'Мега Остатки', 3);

    const api = new OZONAPI(cab);

    // 1) Отчёт по артикулам → offers[]
    let offers = [];
    try {
      const reportId = api.makeArtsReport();
      Utilities.sleep(1000);
      const artsObj = api.getArts(reportId) || {}; // { offerId: CSVcols }
      offers = Object.keys(artsObj);
    } catch (e) {
      console.log('Arts report fail:', cab, (e && e.message) || e);
      offers = [];
    }
    if (!offers.length) {
      console.log('skip cabinet (no offers):', cab);
      continue;
    }

    // 2) offer -> sku (основной путь — getArtsDetails; фолбэк — getSkusByOffers)
    let details = [];
    try { details = api.getArtsDetails(offers) || []; } catch (e) { details = []; }
    const detMap = {};
    for (const it of details) {
      const off = it && (it.offer_id || it.offer);
      if (off) detMap[off] = it;
    }
    // Соберём, что осталось без SKU — дожмём через getSkusByOffers
    const missing = [];
    for (const ofr of offers) {
      const d = detMap[ofr];
      const sku = d && (d.sku || d.sku_id || d.id);
      if (sku == null || sku === '') missing.push(ofr);
    }
    let fallbackMap = {};
    if (missing.length) {
      try { fallbackMap = api.getSkusByOffers(missing) || {}; } catch (e) { fallbackMap = {}; }
    }

    const skuToOffer = {};
    const skus = [];
    for (const ofr of offers) {
      const d = detMap[ofr];
      const sku = (d && (d.sku || d.sku_id || d.id)) || fallbackMap[ofr];
      if (sku != null && sku !== '') {
        const key = String(sku);
        if (!skuToOffer[key]) {
          skuToOffer[key] = ofr;
          skus.push(sku);
        }
      }
    }
    if (!skus.length) {
      console.log('skip cabinet (no skus after mapping):', cab);
      continue;
    }

    // 3) /v1/analytics/stocks по всем SKU (батчи внутри метода)
    let items = [];
    try { items = api.analyticsStocksBySkus(skus) || []; } catch (e) { items = []; }

    // 4) ВЫПЛЕСК: НИЧЕГО не фильтруем — просто добавляем служебные поля
    for (const it of items) {
      const sku = it && (it.sku ?? it.sku_id ?? it.id);
      const offerId = (sku != null ? (skuToOffer[String(sku)] || it.offer_id || it.offer) : (it.offer_id || it.offer)) || '';
      pushFlat(Object.assign({ cabinet: cab, offer_id: offerId, sku: sku }, it));
    }
  }

  // Заголовок: сначала три служебных (если встретились), затем остальное в порядке появления
  const pinned = ['cabinet', 'offer_id', 'sku'];
  const rest = headerOrder.filter(k => pinned.indexOf(k) === -1);
  const header = pinned.filter(k => headerSet.has(k)).concat(rest);

  // Рендер без оформления
  _mega_clearAll_(sh);
  _mega_ensureSize_(sh, (rows.length || 0) + 1, Math.max(1, header.length));
  if (header.length) sh.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) {
    const data = rows.map(r => header.map(h => r[h] == null ? '' : r[h]));
    sh.getRange(2, 1, data.length, header.length).setValues(data);
  }
  SpreadsheetApp.flush();
}

/* ====================== Хелперы (локальные) ====================== */

// Выбранные OZON-кабинеты из ⚙️ Параметры: A..H, D="OZON", H=TRUE
function _mega_getSelectedOZCabinets_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('⚙️ Параметры');
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, 8).getDisplayValues(); // A..H
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const name  = String(vals[i][0] || '').replace(/\u00A0/g, ' ').trim(); // A
    const platf = String(vals[i][3] || '').trim().toUpperCase();           // D
    const on    = String(vals[i][7] || '').trim().toUpperCase();           // H
    const isOn  = (on === 'TRUE' || on === 'ИСТИНА' || on === 'ДА');
    if (name && platf === 'OZON' && isOn) out.push(name);
  }
  return out;
}

// Плоское представление произвольного объекта (вложенности → key_subkey_...)
function _mega_flat_(obj, prefix) {
  const out = {};
  (function go(v, p) {
    if (v == null) { if (p) out[p] = ''; return; }
    else if (Array.isArray(v)) { out[p] = JSON.stringify(v); }
    else if (Object.prototype.toString.call(v) === '[object Date]') { out[p] = v; }
    else if (typeof v === 'object') {
      const ks = Object.keys(v);
      if (!ks.length) { if (p) out[p] = ''; return; }
      for (let i = 0; i < ks.length; i++) {
        const np = p ? (p + '_' + ks[i]) : ks[i];
        go(v[ks[i]], np);
      }
    } else {
      out[p] = v;
    }
  })(obj, prefix || '');
  return out;
}

// Полная очистка листа по значениям (без стилей/форматов)
function _mega_clearAll_(sh) {
  const mr = sh.getMaxRows(), mc = sh.getMaxColumns();
  if (mr > 1) sh.getRange(2, 1, mr - 1, mc).clearContent();
  sh.getRange(1, 1, 1, mc).clearContent();
}

// Гарантируем нужный размер листа (добавляем строки/колонки при необходимости)
function _mega_ensureSize_(sh, needRows, needCols) {
  if (sh.getMaxRows() < needRows) sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());
  if (sh.getMaxColumns() < needCols) sh.insertColumnsAfter(sh.getMaxColumns(), needCols - sh.getMaxColumns());
}
