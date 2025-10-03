/** =========================
 *  WBAPI.gs — минималистичный клиент WB (обновлён)
 * ========================= */
class WB {
  /**
   * @param {string} token - WB API KEY
   */
  constructor(token) {
    this.token = token;
    this.options = {
      muteHttpExceptions: true,
      headers: { Authorization: token },
      method: 'GET',
    };
  }

  /** Универсальный fetch с JSON.parse и экспоненциальным бэкоффом */
  _fetchJson_(url, opt = {}, maxRetries = 6) {
    const baseOptions = this.options || {};
    const options = Object.assign({}, baseOptions, opt, { muteHttpExceptions: true });

    let attempt = 0;
    // eslint-disable-next-line no-constant-condition
    while (true) {
      try {
        const resp = UrlFetchApp.fetch(url, options);
        const code = resp.getResponseCode();
        const text = resp.getContentText('UTF-8');

        if (code >= 200 && code < 300) {
          try { return JSON.parse(text); } catch (_) { return text || null; }
        }
        // 429/5xx — повторяем
        if (code === 429 || (code >= 500 && code < 600)) {
          throw new Error(`HTTP ${code}: ${text}`);
        }
        // прочие коды — как ошибка без повторов
        throw new Error(`HTTP ${code}: ${text}`);
      } catch (err) {
        attempt++;
        if (attempt > maxRetries) {
          throw new Error(`Fetch failed after ${attempt} attempts: ${url}\n${err}`);
        }
        const backoffMs = Math.min(30000, Math.pow(2, attempt) * 200) + Math.floor(Math.random() * 250);
        Utilities.sleep(backoffMs);
      }
    }
  }

  /** ---------- ПУБЛИЧНЫЕ / СЛУЖЕБНЫЕ МЕТОДЫ ---------- */

  /**
   * Карточки (контент)
   * Требуется токен с доступом "Контент".
   * https://content-api.wildberries.ru/content/v2/get/cards/list
   */
  getProducts() {
    const url = 'https://content-api.wildberries.ru/content/v2/get/cards/list';
    const payload = {
      settings: {
        cursor: { limit: 100 },
        filter: { withPhoto: -1 },
      }
    };
    const out = [];
    let attempts = 0;

    // eslint-disable-next-line no-constant-condition
    while (true) {
      try {
        const opt = {
          method: 'POST',
          payload: JSON.stringify(payload),
          contentType: 'application/json',
        };
        const res = this._fetchJson_(url, opt);
        if (!res || !Array.isArray(res.cards)) return out;
        out.push(...res.cards);

        // курсор
        const cur = res.cursor || {};
        if (Number(cur.total || 0) < 100) break;
        payload.settings.cursor.updatedAt = cur.updatedAt;
        payload.settings.cursor.nmID = cur.nmID;
      } catch (err) {
        if (++attempts > 6) throw err;
      }
    }
    return out;
  }

  /**
   * Цены по nmID (чтение)
   * Требуется токен с доступом "Цены и скидки / Аналитика".
   * https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter
   */
  getPrices() {
    const base = 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter';
    const limit = 1000;
    let offset = 0;
    const out = [];
    let attempts = 0;

    // eslint-disable-next-line no-constant-condition
    while (true) {
      try {
        const url = `${base}?limit=${limit}&offset=${offset}`;
        const res = this._fetchJson_(url, { method: 'GET' });
        const list = res && res.data && Array.isArray(res.data.listGoods) ? res.data.listGoods : [];
        out.push(...list);
        if (list.length < limit) break;
        offset += limit;
      } catch (err) {
        if (++attempts > 6) throw err;
      }
    }
    return out;
  }

  /**
   * ОТПРАВКА цен/скидок (update task)
   * Требуется токен с доступом "Цены и скидки".
   *
   * @param {Array<{nmID:number|string, price:number|string, discount?:number|string}>} items
   * @param {{batchSize?:number}} opts
   * @returns {{count:number, uploadIds:string[]}}
   *
   * Эндпоинт: https://discounts-prices-api.wildberries.ru/api/v2/upload/task
   * Формат тела: { data: [ { nmID, price, discount } ] }
   * - nmID: число WB
   * - price: конечная цена (руб)
   * - discount: процент скидки (0..100), можно 0
   */
  setPrices(items, opts = {}) {
    const batchSize = Number(opts.batchSize) || 500;
    if (!Array.isArray(items) || items.length === 0) return { count: 0, uploadIds: [] };

    const endpoint = 'https://discounts-prices-api.wildberries.ru/api/v2/upload/task';
    const uploadIds = [];
    let sent = 0;

    for (let i = 0; i < items.length; i += batchSize) {
      const chunk = items.slice(i, i + batchSize).map(x => ({
        nmID: Number(x.nmID),
        price: Number(x.price),
        discount: Number(x.discount || 0)
      })).filter(x => Number.isFinite(x.nmID) && x.nmID > 0 && Number.isFinite(x.price) && x.price > 0);

      if (!chunk.length) continue;

      const res = this._fetchJson_(endpoint, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({ data: chunk })
      });

      const taskId = res && (res.uploadId || res.taskId || res.id || null);
      if (taskId) uploadIds.push(String(taskId));
      sent += chunk.length;
    }

    return { count: sent, uploadIds };
  }

  /**
   * Заказы (публичная статистика API)
   * https://statistics-api.wildberries.ru/api/v1/supplier/orders?dateFrom=<RFC3339>
   * Требуется токен с правами Statistics.
   */
  getOrders({ dateFrom }) {
    const url = `https://statistics-api.wildberries.ru/api/v1/supplier/orders?dateFrom=${encodeURIComponent(dateFrom)}`;
    return this._fetchJson_(url, { method: 'GET' });
  }

  /**
   * Остатки
   * https://statistics-api.wildberries.ru/api/v1/supplier/stocks?dateFrom=2019-06-20
   * Требуется токен с правами Statistics.
   */
  getStocks() {
    const url = 'https://statistics-api.wildberries.ru/api/v1/supplier/stocks?dateFrom=2019-06-20';
    return this._fetchJson_(url, { method: 'GET' });
  }

  /**
   * Детализация v5 (опционально; 1 req/min на аккаунт)
   * https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod
   */
  getReportDetailByPeriod({ dateFrom, dateTo, limit = 100000 } = {}) {
    if (!dateFrom || !dateTo) {
      throw new Error('getReportDetailByPeriod: нужен dateFrom (RFC3339, Мск) и dateTo (YYYY-MM-DD)');
    }
    const baseUrl = 'https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod';
    let rrdid = 0;
    const out = [];

    // eslint-disable-next-line no-constant-condition
    while (true) {
      const url =
        `${baseUrl}?dateFrom=${encodeURIComponent(dateFrom)}` +
        `&dateTo=${encodeURIComponent(dateTo)}` +
        `&limit=${encodeURIComponent(limit)}` +
        `&rrdid=${encodeURIComponent(rrdid)}`;

      const data = this._fetchJson_(url, { method: 'GET' });
      if (!Array.isArray(data) || data.length === 0) break;

      out.push(...data);
      const last = data[data.length - 1];
      if (!last || typeof last.rrd_id !== 'number') break;
      if (data.length < limit) break;
      rrdid = last.rrd_id;

      // Ограничение WB: 1 запрос/мин
      Utilities.sleep(61000);
    }
    return out;
  }

  /**
   * Рейтинг/отзывы — публичный cards/v2 (без токена)
   * https://card.wb.ru/cards/v2/detail?nm=<ids>
   */
  static getRnRPublic(nmIds) {
    if (!Array.isArray(nmIds) || nmIds.length === 0) return [];
    const page = 50;
    const out = [];

    for (let i = 0; i < nmIds.length; i += page) {
      const chunk = nmIds.slice(i, i + page);
      const url = `https://card.wb.ru/cards/v2/detail?appType=1&curr=rub&dest=123589785&lang=ru&nm=${chunk.join(';')}`;
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json = JSON.parse(resp.getContentText('UTF-8'));
      if (json && json.data && Array.isArray(json.data.products)) {
        out.push(...json.data.products);
      }
    }
    return out;
  }

  /**
   * Комиссии по категориям (common-api)
   * https://common-api.wildberries.ru/api/v1/tariffs/commission?locale=ru
   */
  getCategoryCommissions() {
    const url = 'https://common-api.wildberries.ru/api/v1/tariffs/commission?locale=ru';
    const res = this._fetchJson_(url, { method: 'GET' });
    if (res && Array.isArray(res.report)) return res.report;
    return [];
  }

  /** Утилита: безопасный расчёт объёма (Д*Ш*В / 1000) */
  static calcVolumeLiters(card) {
    const d = card && card.dimensions;
    if (!d) return '';
    const L = Number(d.length || d.L || 0);
    const W = Number(d.width  || d.W || 0);
    const H = Number(d.height || d.H || 0);
    if (!(L && W && H)) return '';
    return (L * W * H) / 1000; // литры/дм3
  }

  /** Утилита: извлечь человеко-читаемую категорию и её ID */
  static pickSubject(card) {
    const id = card.subjectId || card.subjectID || null;
    const name = card.subjectName || card.subject_name || '';
    return { subjectId: id, subjectName: name };
  }

  /** Утилита: выжать цену из объекта цен */
  static pickPrice(priceObj) {
    if (!priceObj) return '';
    const s0 = (priceObj.sizes && priceObj.sizes[0]) || {};
    // Приоритет: price -> discountedPrice -> basicSalePrice -> basicPrice
    return (
      s0.price ??
      s0.discountedPrice ??
      s0.basicSalePrice ??
      s0.basicPrice ??
      ''
    );
  }

  /** Утилита: выжать комиссии (пытаемся распознать разные ключи API) */
  static pickCommissions(rec) {
    const fbo  = rec.kgvpMarketplace ?? rec.fbo ?? rec.fboKgvp ?? rec.fboPercent ?? '';
    const fbs  = rec.kgvpSupplier   ?? rec.fbs ?? rec.fbsKgvp ?? rec.fbsPercent ?? '';
    const rfbs = rec.kgvpBooking    ?? rec.rfbs ?? rec.booking ?? rec.rfbsPercent ?? '';
    return {
      FBO:  fbo  === '' ? '' : Number(fbo),
      FBS:  fbs  === '' ? '' : Number(fbs),
      RFBS: rfbs === '' ? '' : Number(rfbs),
    };
  }
}
