// Универсальный клиент Ozon Seller API с инициализацией ключей из листа "⚙️ Параметры"
var OZONAPI = (function () {
  // ===== Конструктор =====
  function OZONAPI(cabinetName, storeId) {
    if (!OZONAPI._accounts) OZONAPI.initFromSheet();
    if (!OZONAPI._accounts[cabinetName]) {
      throw new Error('Кабинет не найден: ' + cabinetName);
    }

    var acc = OZONAPI._accounts[cabinetName];
    this.url = 'https://api-seller.ozon.ru';
    this.headers = {
      'Client-Id': String(acc.client_id),
      'Api-Key': acc.api_key,
      'Content-Type': 'application/json; charset=utf-8'
    };
    this.storeId = storeId || null;

    // Локальный кэш дерева категорий (на инстанс)
    this._categoryCache = null;
  }

  // ====== Статические поля ======
  OZONAPI._accounts = null;               // { [cabinetName]: { client_id, api_key, name, short } }
  OZONAPI._categoryCacheStatic = null;    // общий кэш дерева категорий на время запуска

  // ====== Статика: чтение аккаунтов из "⚙️ Параметры" ======
  OZONAPI.initFromSheet = function () {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('⚙️ Параметры');
    if (!sh) throw new Error('Нет листа "⚙️ Параметры"');

    // Структура: A=Кабинет, B=ID кабинета (Client-Id), C=API KEY, D=Площадка/Тип (должно быть "OZON")
    var rows = sh.getRange('A2:D').getValues();

    OZONAPI._accounts = rows.reduce(function (acc, row) {
      var cabinetName = String(row[0] || '').trim();
      var clientId    = row[1];
      var apiKey      = row[2];
      var platform    = String(row[3] || '').trim().toUpperCase(); // D

      if (cabinetName && clientId && apiKey && platform === 'OZON') {
        acc[cabinetName] = {
          client_id: clientId,
          api_key: apiKey,
          name: cabinetName,
          short: ''
        };
      }
      return acc;
    }, {});
  };

  OZONAPI.getAccounts = function () {
    if (!OZONAPI._accounts) OZONAPI.initFromSheet();
    return OZONAPI._accounts;
  };

  // ====== Вспомогательный POST с JSON, ретраями и бэкоффом ======
  // opts = {payload, maxAttempts, baseDelayMs}
  OZONAPI.prototype._post = function (endpoint, opts) {
    var url = (endpoint.indexOf('http') === 0) ? endpoint : (this.url + endpoint);
    var payload = (opts && opts.payload) || {};
    var maxAttempts = (opts && opts.maxAttempts) || 10;
    var baseDelay = (opts && opts.baseDelayMs) || 300;

    for (var attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        var resp = UrlFetchApp.fetch(url, {
          method: 'POST',
          muteHttpExceptions: true,
          headers: this.headers,
          payload: JSON.stringify(payload),
        });

        var code = resp.getResponseCode();
        var text = resp.getContentText();

        var data;
        try { data = JSON.parse(text || '{}'); } catch (parseErr) {
          if (code >= 500 || code === 429) {
            Utilities.sleep(baseDelay * Math.pow(1.8, attempt - 1));
            continue;
          }
          throw new Error('Bad response (' + code + '), not JSON: ' + String(text).slice(0, 200));
        }

        if (code >= 200 && code < 300) return data;

        if (code === 429 || code >= 500) {
          Utilities.sleep(baseDelay * Math.pow(1.8, attempt - 1));
          continue;
        }

        throw new Error('HTTP ' + code + ' ' + (data && (data.message || data.error)) );

      } catch (err) {
        if (attempt >= maxAttempts) throw err;
        Utilities.sleep(baseDelay * Math.pow(1.8, attempt - 1));
      }
    }
    throw new Error('Не удалось выполнить запрос: ' + endpoint);
  };

  // ====== Отчёты/артикулы ======
  OZONAPI.prototype.makeArtsReport = function () {
    var data = this._post('/v1/report/products/create', {
      payload: { language: 'DEFAULT', visibility: 'ALL' },
      maxAttempts: 10,
      baseDelayMs: 300
    });
    if (data && data.result && data.result.code) return data.result.code;
    throw new Error('Не удалось сгенерировать отчёт по артикулам');
  };

  OZONAPI.prototype.getReportInfo = function (reportId) {
    var attempts = 0;
    while (true) {
      attempts++;
      var data = this._post('/v1/report/info', {
        payload: { code: reportId },
        maxAttempts: 3,
        baseDelayMs: 300
      });
      if (data && data.result) {
        if (data.result.status === 'success') return data.result.file;
        if (attempts > 20) throw new Error('Не дождались отчёта (status=' + data.result.status + ')');
        Utilities.sleep(1000);
        continue;
      }
      if (attempts > 20) throw new Error('Не дождались отчёта');
      Utilities.sleep(1000);
    }
  };

  // === CSV с авто-детектом кодировки и «;»
  OZONAPI.prototype.getCSV = function (link) {
    if (!link) throw new Error('Нет ссылки на файл');
    var resp = UrlFetchApp.fetch(link);
    var blob = resp.getBlob();

    // Попытка 1: авто
    try {
      return Utilities.parseCsv(blob.getDataAsString(), ';');
    } catch (_) {}

    // Попытка 2: CP1251
    try {
      return Utilities.parseCsv(blob.getDataAsString('Windows-1251'), ';');
    } catch (_) {}

    // Попытка 3: UTF-8
    var txt = resp.getContentText('UTF-8');
    return Utilities.parseCsv(txt, ';');
  };

  OZONAPI.prototype.getXLSX = function (link) {
    if (!link) throw new Error('Нет ссылки на файл');
    var file = UrlFetchApp.fetch(link);
    var tempId = Drive.Files.insert({ title: 'tempFile' }, file.getBlob(), { convert: true }).getId();
    var values = SpreadsheetApp.openById(tempId).getDataRange().getValues();
    Drive.Files.remove(tempId);
    return values;
  };

  OZONAPI.prototype.getArts = function (reportId) {
    reportId = reportId || this.makeArtsReport();
    var link = this.getReportInfo(reportId);
    var fileData = this.getCSV(link);
    var header = fileData[0] || [];
    var headerMap = header.reduce(function (acc, col, idx) { acc[col] = idx; return acc; }, {});
    return fileData.slice(1).reduce(function (acc, row) {
      var idx = headerMap['Артикул'];
      var art = (idx != null ? String(row[idx] || '') : '').replace(/\'/g, '');
      if (!art) return acc;
      acc[art] = {};
      for (var col in headerMap) {
        var v = row[headerMap[col]];
        acc[art][col] = (v == null ? '' : String(v)).replace(/\'/g, '');
      }
      return acc;
    }, {});
  };

  OZONAPI.prototype.getArtsDetails = function (arts) {
    arts = arts || [];
    var url = '/v3/product/info/list';
    var limit = 1000;
    var finArr = [];
    for (var i = 0; i < arts.length; i += limit) {
      var slice = arts.slice(i, i + limit);
      var data = this._post(url, {
        payload: { offer_id: slice },
        maxAttempts: 10,
        baseDelayMs: 300
      });
      if (data && data.items) finArr = finArr.concat(data.items);
    }
    return finArr;
  };

  // ====== Дерево категорий/типов (RU) — общий статический кэш ======
  OZONAPI.prototype.getCategoryTreeRU = function () {
    if (OZONAPI._categoryCacheStatic) return OZONAPI._categoryCacheStatic;
    if (this._categoryCache) return this._categoryCache;

    var data = this._post('/v1/description-category/tree', {
      payload: { language: 'RU' },
      maxAttempts: 10,
      baseDelayMs: 500
    });

    var byCategoryId = {};
    var byTypeId     = {};
    var typeNameByTypeId = {};

    var walk = function (node, parentCategoryName) {
      if (!node) return;
      var catId   = node.description_category_id;
      var catName = node.category_name || parentCategoryName || '';

      if (catId && node.category_name) byCategoryId[String(catId)] = node.category_name;

      if (node.type_id) {
        var tid = String(node.type_id);
        byTypeId[tid] = catName || '';
        if (node.type_name) typeNameByTypeId[tid] = node.type_name;
      }

      if (Array.isArray(node.children)) {
        node.children.forEach(function (child) { walk(child, node.category_name || parentCategoryName || ''); });
      }
    };

    if (data && Array.isArray(data.result)) {
      data.result.forEach(function (root) { walk(root, null); });
    }

    var cache = { byCategoryId: byCategoryId, byTypeId: byTypeId, typeNameByTypeId: typeNameByTypeId };
    this._categoryCache = cache;
    OZONAPI._categoryCacheStatic = cache;
    return cache;
  };

  // ====== Цены ======
  OZONAPI.prototype.getPrices = function () {
    var url = '/v5/product/info/prices';
    var payload = { cursor: '', limit: 1000, filter: { visibility: 'ALL' } };
    var finArr = [];
    var guard = 0;

    while (true) {
      guard++;
      var data = this._post(url, {
        payload: payload,
        maxAttempts: 10,
        baseDelayMs: 300
      });

      if (data && data.items) finArr = finArr.concat(data.items);

      var hasCursor = !!(data && data.cursor);
      if (!hasCursor || !data.items || data.items.length < payload.limit || guard > 200) break;

      payload.cursor = data.cursor;
    }
    return finArr;
  };

// Точечное получение цен по списку offer_id (быстро, без CSV)
OZONAPI.prototype.getPricesByOffers = function (offerIds) {
  if (!Array.isArray(offerIds) || offerIds.length === 0) return {};
  var url = '/v5/product/info/prices';

  function toNum_(v) {
    if (typeof REF !== 'undefined' && typeof REF.toNumber === 'function') return REF.toNumber(v);
    var s = String(v == null ? '' : v).replace(/\s/g, '').replace(',', '.');
    var n = parseFloat(s);
    return isFinite(n) ? n : NaN;
  }

  var CH = 1000; // безопасный батч
  var map = {};

  for (var i = 0; i < offerIds.length; i += CH) {
    var slice = offerIds.slice(i, i + CH);

    var data = this._post(url, {
      payload: { cursor: "", limit: CH, filter: { visibility: 'ALL', offer_id: slice } },
      maxAttempts: 10,
      baseDelayMs: 300
    });

    var items = (data && data.items) || [];
    for (var j = 0; j < items.length; j++) {
      var it = items[j] || {};
      var offer = String(it.offer_id || it.offer || '').trim();
      if (!offer) continue;

      // пытаемся вытащить числовую цену из разных полей, Ozon менял схему
      var p = null;
      if (it.price && it.price.price != null) p = it.price.price;
      if (p == null && it.price && it.price.min_price != null) p = it.price.min_price;
      if (p == null && it.current_prices && it.current_prices.price != null) p = it.current_prices.price;

      var num = toNum_(p);
      if (!isNaN(num)) map[offer] = num;
    }
  }
  return map;
};




  // ====== Остатки (FBS отчёт) ======
  OZONAPI.prototype.makeFbsReport = function () {
    var payload = { language: 'DEFAULT' };
    if (this.storeId != null && this.storeId !== '') {
      payload.warehouse_id = [this.storeId];
    }
    var data = this._post('/v1/report/warehouse/stock', {
      payload: payload,
      maxAttempts: 10,
      baseDelayMs: 300
    });
    if (data && data.result && data.result.code) return data.result.code;
    throw new Error('Не удалось сгенерировать FBS отчёт');
  };

  OZONAPI.prototype.getFbsStocks = function () {
    var reportId = this.makeFbsReport();
    var link = this.getReportInfo(reportId);
    var values = this.getXLSX(link);
    var header = values[0] || [];
    var headerMap = header.reduce(function (acc, col, idx) { acc[idx] = col; return acc; }, {});
    return values.slice(1).map(function (row) {
      return row.reduce(function (acc, val, idx) {
        acc[headerMap[idx]] = val;
        return acc;
      }, {});
    });
  };

  // ====== Запись остатков/цен ======
  OZONAPI.prototype.setStocks = function (stocksArr) {
    var url = '/v2/products/stocks';
    var partSize = 100;
    var fin = [];
    for (var i = 0; i < stocksArr.length; i += partSize) {
      var chunk = stocksArr.slice(i, i + partSize);
      var data = this._post(url, {
        payload: { stocks: chunk },
        maxAttempts: 10,
        baseDelayMs: 300
      });
      if (data && data.result) fin = fin.concat(data.result);
    }
    return fin;
  };

  OZONAPI.prototype.setPrices = function (pricesArr) {
    var url = '/v1/product/import/prices';
    var partSize = 1000;
    var fin = [];
    for (var i = 0; i < pricesArr.length; i += partSize) {
      var chunk = pricesArr.slice(i, i + partSize);
      var data = this._post(url, {
        payload: { prices: chunk },
        maxAttempts: 10,
        baseDelayMs: 300
      });
      if (data && data.result) fin = fin.concat(data.result);
    }
    return fin;
  };

  // ====== Заказы FBS ======
  OZONAPI.prototype.getOrdersFbs = function (days) {
    days = days == null ? 120 : days;
    var url = '/v3/posting/fbs/list';

    var now = new Date();
    now.setDate(now.getDate() + 1);
    now.setHours(3, 0, 0, 0);
    now.setMilliseconds(now.getMilliseconds() - 1);
    var to = now.toISOString();

    now.setDate(now.getDate() - days - 1);
    now.setMilliseconds(now.getMilliseconds() + 1);
    var since = now.toISOString();

    var payload = {
      dir: 'ASC',
      limit: 1000,
      offset: 0,
      filter: { since: since, to: to },
      with: { analytics_data: true, barcodes: true, financial_data: true, translit: true }
    };

    var fin = [];
    var guard = 0;

    while (true) {
      guard++;
      var data = this._post(url, {
        payload: payload,
        maxAttempts: 10,
        baseDelayMs: 300
      });

      if (data && data.result && Array.isArray(data.result.postings)) {
        fin = fin.concat(data.result.postings);
        if (data.result.postings.length < payload.limit || guard > 500) break;

        payload.offset += payload.limit;
        if (payload.offset >= 20000) {
          var last = data.result.postings[data.result.postings.length - 1];
          payload.filter.since = new Date(last.in_process_at).toISOString();
          payload.offset = 0;
        }
      } else {
        break;
      }
    }
    return fin;
  };

  // ====== Склады продавца ======
  OZONAPI.prototype.getWarehouses = function () {
    var data = this._post('/v1/warehouse/list', {
      payload: {},
      maxAttempts: 10,
      baseDelayMs: 300
    });
    return (data && Array.isArray(data.result)) ? data.result : [];
  };

  // ====== Заказы FBO ======
  OZONAPI.prototype.getOrdersFbo = function (days) {
    days = days == null ? 60 : days;
    var url = '/v2/posting/fbo/list';

    var to = new Date();
    to.setDate(to.getDate() - 1);
    to.setHours(23, 59, 59, 999);

    var since = new Date(to);
    since.setDate(since.getDate() - (days - 1));
    since.setHours(0, 0, 0, 0);

    var payload = {
      dir: 'ASC',
      limit: 1000,
      offset: 0,
      filter: { since: since.toISOString(), to: to.toISOString() },
      with: { analytics_data: true, barcodes: true, financial_data: true, translit: true }
    };

    var fin = [];
    var guard = 0;

    while (true) {
      guard++;
      var data = this._post(url, {
        payload: payload,
        maxAttempts: 6,
        baseDelayMs: 250
      });

      if (data && Array.isArray(data.result)) {
        fin = fin.concat(data.result);
        if (data.result.length < payload.limit || guard > 200) break;

        payload.offset += payload.limit;
        if (payload.offset >= 20000) {
          var last = data.result[data.result.length - 1];
          payload.filter.since = new Date(last.in_process_at || last.created_at).toISOString();
          payload.offset = 0;
        }
      } else {
        break;
      }
    }
    return fin;
  };

  // ====== Маппинг offer_id -> sku (батчами по 1000) ======
  OZONAPI.prototype.getSkusByOffers = function(offerIds) {
    if (!Array.isArray(offerIds) || offerIds.length === 0) return {};
    var url = '/v3/product/info/list';
    var limit = 1000;
    var map = {};
    for (var i = 0; i < offerIds.length; i += limit) {
      var slice = offerIds.slice(i, i + limit);
      var data = this._post(url, {
        payload: { offer_id: slice },
        maxAttempts: 6,
        baseDelayMs: 250
      });
      var items = (data && data.items) || [];
      for (var j = 0; j < items.length; j++) {
        var it = items[j] || {};
        var offer = it.offer_id || it.offer || '';
        var sku   = it.sku || it.sku_id || it.id || '';
        if (offer) map[offer] = sku;
      }
    }
    return map;
  };

  // ====== Остатки по SKU через analytics/stocks (батчи по 100) ======
  OZONAPI.prototype.analyticsStocksBySkus = function(skus) {
    if (!Array.isArray(skus) || skus.length === 0) return [];
    var url = '/v1/analytics/stocks';
    var out = [];
    var batch = 100;
    for (var i = 0; i < skus.length; i += batch) {
      var part = skus.slice(i, i + batch);
      var data = this._post(url, {
        payload: { skus: part },
        maxAttempts: 6,
        baseDelayMs: 250
      });
      var items = (data && data.items) || [];
      for (var k = 0; k < items.length; k++) out.push(items[k]);
      Utilities.sleep(200);
    }
    return out;
  };

  return OZONAPI;
})();
