/** ===============================
 *  Import_Sklad.gs — автономный сборщик
 *  ===============================
 * Делает:
 *  1) Читает токен МС из 🍔 СС!AG2
 *  2) Тянет ассортимент из МойСклад (Код/Производитель/Модель/Вес/Выключен)
 *  3) Выгружает склады в 🍔 СС!AI:AL (Склад/Код/Доступно/Ожидание)
 *  4) Собирает «Приёмки» (внешний файл), берёт Очередность=1
 *  5) Считает комплекты (O:Q):
 *      - СС в валюте (юань) = Σ(СС(состав, юань) * кол-во)
 *      - Наличие = min_i floor(наличие(part_i)/qty_i)   ← ТОЛЬКО ДЛЯ «Наличие»
 *      - «В пути», «В поставке OZ», «В поставке WB» — только прямые (без комплектов)
 *  6) Считает «СС+Упак+Дост»:
 *      - СС (руб) + Упаковка (руб, из Y) + Доставка(вес_кг * тариф$ * курс$ *1.1)
 *        (вес берём в граммах → кг; тариф «доставка» и «доллар» из L:M)
 *  7) Пишет итог в 🍔 СС!A:J
 *       A:Товар B:Производитель C:Модель D:СС в валюте E:Валюта
 *       F:СС+Упак+Дост G:Наличие H:В пути I:В поставке J:Не закупается
 */

const RECIEVES_SPREADSHEET_ID = '1wX4N41BDVBEJ4UUOdO2bZAhYZG7TaJuOMReI6g473aE';

function Import_Sklad() {
  const ss   = SpreadsheetApp.getActive();
  const shCC = mustSheet(ss, '🍔 СС');

  const T0 = Date.now();
  const t = (label) => `[+${String(Date.now()-T0).padStart(6,' ')} ms] ${label}`;
  console.log(t('START Import_Sklad'));
  ss.toast('Импорт МойСклад → расчёт СС', 'Склад + СС', 3);

  try {
    // === 0) Токен МС
    const token = String(shCC.getRange('AG2').getValue() || '').trim();
    if (!token) throw new Error('Пустой токен МС в 🍔 СС!AG2');

    // === 1) Ассортимент из МС
    const ms  = new MoySklad(token);
    const prods = fetchProductsFromMS_(ms); // [{code, manufacturer, model, weightRaw, disabled}]
    const prodByCode = {};
    for (const p of prods) {
      prodByCode[p.code] = {
        manufacturer: p.manufacturer || '',
        model: p.model || '',
        weightRaw: ('weightRaw' in p ? p.weightRaw : ''),
        disabled: !!p.disabled
      };
    }

    // === 2) Выгрузка складов → 🍔 СС!AI:AL
    const stockRows = exportStocksToCC_(ms, shCC); // [ [store, code, available, inTransit], ... ]
    let stockAgg  = aggregateStocks_(stockRows);    // code -> {availMain, transitMain, vpostOZ, vpostWB}

    // === 3) Курсы L:M (юань/доллар/доставка)
    const lastRowCC = shCC.getLastRow();
    const rates = {};
    if (lastRowCC >= 1) {
      const lm = shCC.getRange(1, 12, lastRowCC, 2).getValues(); // L(12):M(13)
      for (const [name, val] of lm) {
        const k = String(name || '').trim().toLowerCase();
        const v = toNum(val);
        if (!k || !isFinite(v)) continue;
        if (k === 'юань' || k === 'доллар' || k === 'доставка') rates[k] = v;
      }
    }

    // === 4) Упаковка S:Y
    const packByCode = {};
    if (lastRowCC >= 2) {
      const pack = shCC.getRange(2, 19, lastRowCC - 1, 7).getValues(); // S..Y
      for (const row of pack) {
        const code = String(row[0] || '').trim(); // S
        if (!code) continue;
        const add = toNum(row[6]);                // Y
        packByCode[code] = isFinite(add) ? add : 0;
      }
    }

    // === 5) Комплекты O:Q
    const kits = readKits_(shCC); // {kit: [{part, qty}, ...]}

    // === 6) Приёмки (внешний файл), берём Очередность=1
    const ext   = SpreadsheetApp.openById(RECIEVES_SPREADSHEET_ID);
    const shRec = mustSheet(ext, 'Приёмки');
    const rec   = shRec.getDataRange().getValues();
    const rHdr  = headerMap(rec[0] || []);
    const priceCol = ('СС в валюте' in rHdr) ? 'СС в валюте'
                     : ('СС в валюте документа' in rHdr) ? 'СС в валюте документа'
                     : null;
    if (!priceCol) throw new Error('В "Приёмки" нет "СС в валюте" или "СС в валюте документа"');
    ['Товар','Валюта','Очередность'].forEach(c => mustHave(rHdr, c, 'Приёмки'));

    const currMap = {
      'доллар сша':       'доллар',
      'китайский юань':   'юань',
      'российский рубль': 'рубль'
    };

    // code -> {costDoc, curr}
    const priceByCode = {};
    for (const row of rec.slice(1)) {
      const ord = toNum(row[rHdr['Очередность']]);
      if (ord !== 1) continue;
      const code = String(row[rHdr['Товар']] || '').trim();
      if (!code) continue;
      const cost = toNum(row[rHdr[priceCol]]);
      const msC  = String(row[rHdr['Валюта']] || '').trim().toLowerCase();
      const cur  = currMap[msC] || (msC || '');
      priceByCode[code] = { costDoc: isFinite(cost) ? cost : '', curr: cur };
    }

    // === 7) Набор всех кодов
    const codeSet = new Set();
    prods.forEach(p => codeSet.add(p.code));
    Object.keys(priceByCode).forEach(c => codeSet.add(c));
    Object.keys(stockAgg).forEach(c => codeSet.add(c));
    Object.keys(kits).forEach(k => codeSet.add(k));
    Object.values(kits).forEach(arr => arr.forEach(({part}) => codeSet.add(part)));

    // === 8) Комплектные СС в юанях (если все составные в юанях)
    for (const kit of Object.keys(kits)) {
      if (!codeSet.has(kit)) continue;
      let sumYuan = 0;
      let ok = true;
      for (const { part, qty } of kits[kit]) {
        const info = priceByCode[part];
        const cost = info ? info.costDoc : '';
        const cur  = info ? (info.curr || '') : '';
        if (cost === '' || cur !== 'юань') { ok = false; break; }
        sumYuan += Number(cost) * Number(qty || 0);
      }
      if (ok) priceByCode[kit] = { costDoc: sumYuan, curr: 'юань' };
    }

    // === 9) Комплектное «Наличие» (ТОЛЬКО availMain!)
    // «В пути», «В поставке OZ/WB» — ТОЛЬКО прямые значения из AI:AL.
    stockAgg = applyKitAvailOnly_(kits, stockAgg);

    // === 10) Сборка итоговых строк A:J
    const HEADER = [
      'Товар','Производитель','Модель','СС в валюте','Валюта',
      'СС+Упак+Дост','Наличие','В пути','В поставке','Не закупается'
    ];

    const codes = Array.from(codeSet).filter(Boolean)
      .sort((a,b)=>String(a).localeCompare(String(b)));
    const out = [];
    for (const code of codes) {
      const p = prodByCode[code] || {manufacturer:'', model:'', weightRaw:'', disabled:false};
      const r = priceByCode[code] || {costDoc:'', curr:''};
      const s = stockAgg[code]    || {availMain:0, transitMain:0, vpostOZ:0, vpostWB:0};

      const curr = String(r.curr || '').toLowerCase();
      let costRub = '';
      if (r.costDoc !== '' && isFinite(Number(r.costDoc))) {
        if (curr === 'рубль' || curr === 'rub' || curr === 'rur') {
          costRub = Number(r.costDoc);
        } else if (curr === 'юань') {
          const rate = rates['юань']; if (isFinite(rate)) costRub = Number(r.costDoc) * rate;
        } else if (curr === 'доллар') {
          const rate = rates['доллар']; if (isFinite(rate)) costRub = Number(r.costDoc) * rate;
        }
      }

      const packAdd = isFinite(toNum(packByCode[code])) ? Number(packByCode[code]) : 0;

      const weightKg = isFinite(toNum(p.weightRaw)) ? (Number(p.weightRaw) / 1000) : 0;
      const rateDeliveryUSD = isFinite(toNum(rates['доставка'])) ? Number(rates['доставка']) : 0;
      const deliveryUSD = weightKg * rateDeliveryUSD;
      const usdRate = isFinite(toNum(rates['доллар'])) ? Number(rates['доллар']) : 0;
      const deliveryRubFinal = (deliveryUSD * usdRate) * 1.1;

      const baseRub = isFinite(toNum(costRub)) ? Number(costRub) : 0;
      const totalRub = baseRub + packAdd + deliveryRubFinal;

      // Новая логика колонок I и J:
      const vPostavke = Number(s.vpostOZ || 0) + Number(s.vpostWB || 0); // I: «В поставке»
const notPurchasingText = p.disabled ? 'да' : '';                                // J: «Не закупается» (из «Выключен»)

      out.push([
        code, p.manufacturer, p.model,
        r.costDoc === '' ? '' : Number(r.costDoc),
        curr || '',
        Number(totalRub),
        s.availMain, s.transitMain, vPostavke, notPurchasingText
      ]);
    }

    // === 11) Запись в 🍔 СС!A:J
    clearBlock(shCC, 1, 1, shCC.getMaxRows(), 10);
    ensureCols(shCC, 10);
    shCC.getRange(1,1,1,10).setValues([HEADER]).setFontWeight('bold');
    if (out.length) shCC.getRange(2,1,out.length,10).setValues(out);

    shCC.setFrozenRows(1);
    if (out.length) {
      shCC.getRange(2,4,out.length,1).setNumberFormat('#,##0.00'); // D
      shCC.getRange(2,6,out.length,1).setNumberFormat('#,##0.00'); // F
      shCC.getRange(2,7,out.length,3).setNumberFormat('#,##0');    // G:H:I (только числа)
      // J: «Не закупается» — логическое; формат не задаём
    }

    // === 12) Лог в «Обновления»
    safeLogRun_MS_(['МойСклад']);

    ss.toast('Склад + СС: обновлено', 'Готово', 3);
    console.log(t(`END Import_Sklad | rows=${out.length}`));
  } catch (e) {
    safeLogRun_MS_([]);
    ss.toast('Склад + СС: ошибка, см. журнал', 6);
    console.error(t(`ERROR Import_Sklad: ${e && e.stack || e}`));
    throw e;
  }
}

function safeLogRun_MS_(cabs) {
  try {
    if (typeof REF !== 'undefined' && typeof REF.logRun === 'function') {
      REF.logRun('Склад + СС', Array.isArray(cabs) ? cabs : ['МойСклад'], 'MOYSKLAD');
    }
  } catch (_) {}
}

/* ================== Вспомогательные блоки ================== */

// --- МС: ассортимент (code, manufacturer, model, weightRaw, disabled[«Выключен»])
function fetchProductsFromMS_(ms) {
  const fin = [];
  const url = 'https://api.moysklad.ru/api/remap/1.2/entity/assortment?extend=attributes';
  const limit = 1000;
  let offset = 0, tries = 0;

  while (true) {
    const full = `${url}&limit=${limit}&offset=${offset}`;
    try {
      const resp = ms.fetch(full, { method: 'GET' });
      const data = JSON.parse(resp.getContentText());
      const rows = data.rows || [];
      for (const item of rows) {
        const code = item && item.code ? String(item.code) : '';
        if (!code) continue;
        let manufacturer = '', model = '', weightRaw = '', disabled = false;

        const attrs = Array.isArray(item.attributes) ? item.attributes : [];
        for (const a of attrs) {
          if (!a || a.name == null) continue;
          if (a.name === 'Производитель' && 'value' in a) manufacturer = a.value;
          else if (a.name === 'Модель' && 'value' in a)   model = a.value;
          else if (a.name === 'Вес' && 'value' in a)      weightRaw = a.value; // граммы (как есть)
          else if (a.name === 'Выключен' && 'value' in a) disabled = toBool(a.value);
        }
        if (!weightRaw && item.weight != null) weightRaw = item.weight; // fallback

        fin.push({ code, manufacturer, model, weightRaw, disabled });
      }
      if (rows.length < limit) break;
      offset += limit;
      tries = 0;
    } catch (e) {
      tries++;
      if (tries > 10) throw e;
      Utilities.sleep(tries * tries * 1000);
    }
  }
  return fin;
}

// --- Выгрузка складов в 🍔 СС!AI:AL, и вернуть «сырые» строки
function exportStocksToCC_(ms, shCC) {
  // Заголовки AI:AL
  shCC.getRange('AI1:AL1').setValues([['Склад','Код','Доступно','Ожидание']]);

  const stores = getStores_(ms); // [{name, meta:{href}}]
  const finObj = {}; // storeName -> code -> {stock, inTransit}

  for (const store of stores) {
    const storeHref = store?.meta?.href;
    if (!storeHref) continue;
    const items = getAssortmentByStore_(ms, storeHref); // rows with .code, .stock, .inTransit
    for (const it of items) {
      const code = it && it.code ? String(it.code) : '';
      if (!code) continue;
      if (!finObj[store.name]) finObj[store.name] = {};
      if (!finObj[store.name][code]) finObj[store.name][code] = { stock: 0, inTransit: 0 };
      finObj[store.name][code].stock     = it.stock || 0;
      finObj[store.name][code].inTransit = it.inTransit || 0;
    }
  }

  // В плоский массив
  const rows = [];
  Object.keys(finObj).forEach(store => {
    Object.keys(finObj[store]).forEach(code => {
      const v = finObj[store][code];
      rows.push([store, code, v.stock || 0, v.inTransit || 0]);
    });
  });

  // Сортировка по коду
  rows.sort((a,b) => String(a[1]).localeCompare(String(b[1])));

  // Очистка и запись начиная со второй строки
  const maxRows = shCC.getMaxRows();
  if (maxRows > 1) shCC.getRange(2, 35, maxRows - 1, 4).clearContent(); // AI=35
  if (rows.length) shCC.getRange(2, 35, rows.length, 4).setValues(rows);

  return rows; // [ [store, code, available, inTransit], ... ]
}

// --- Агрегаты по складам из сырья AI:AL
function aggregateStocks_(rows) {
  const MAIN = 'Основной склад';
  const POZ  = 'В поставке OZ';
  const PWB  = 'В поставке WB';
  const agg = {}; // code -> {availMain, transitMain, vpostOZ, vpostWB}
  for (const [store, code, avail, wait] of rows) {
    if (!agg[code]) agg[code] = { availMain:0, transitMain:0, vpostOZ:0, vpostWB:0 };
    if (store === MAIN) {
      agg[code].availMain   += Number(avail) || 0;
      agg[code].transitMain += Number(wait)  || 0;
    } else if (store === POZ) {
      agg[code].vpostOZ     += Number(avail) || 0; // ТОЛЬКО «Доступно»
    } else if (store === PWB) {
      agg[code].vpostWB     += Number(avail) || 0; // ТОЛЬКО «Доступно»
    }
  }
  return agg;
}

// --- Комплекты O:Q → {kit: [{part, qty}, ...]}
function readKits_(shCC) {
  const last = shCC.getLastRow();
  const res = {};
  if (last < 2) return res;
  const oq = shCC.getRange(2, 15, last - 1, 3).getValues(); // O(15):Q(17)
  for (const [kitRaw, partRaw, qtyRaw] of oq) {
    const kit  = String(kitRaw  || '').trim();
    const part = String(partRaw || '').trim();
    const qty  = toNum(qtyRaw);
    if (!kit || !part || !isFinite(qty)) continue;
    if (!res[kit]) res[kit] = [];
    res[kit].push({ part, qty: Number(qty) });
  }
  return res;
}

// --- Применить комплектность ТОЛЬКО к «Наличие» (availMain)
function applyKitAvailOnly_(kits, stockAgg) {
  const out = Object.assign({}, stockAgg);
  for (const kit of Object.keys(kits)) {
    const parts = kits[kit];
    if (!parts || !parts.length) continue;
    let potAvail = Infinity;

    for (const {part, qty} of parts) {
      const s = out[part] || {availMain:0};
      const q = Number(qty) || 0;
      if (q <= 0) { potAvail = 0; break; }
      const partAvail = Math.floor((Number(s.availMain) || 0) / q);
      potAvail = Math.min(potAvail, partAvail);
    }

    if (!isFinite(potAvail)) potAvail = 0;
    if (!out[kit]) out[kit] = { availMain:0, transitMain:0, vpostOZ:0, vpostWB:0 };
    out[kit].availMain = Math.max(0, potAvail | 0);
  }
  return out;
}

// --- МС helpers (stores & assortment by store)
function getStores_(ms) {
  const url = 'https://api.moysklad.ru/api/remap/1.2/entity/store';
  const resp = ms.fetch(url, { method: 'GET' });
  const data = JSON.parse(resp.getContentText());
  return data.rows || [];
}
function getAssortmentByStore_(ms, storeHref) {
  const fin = [];
  const url = 'https://api.moysklad.ru/api/remap/1.2/entity/assortment?extend=attributes';
  const limit = 1000;
  let offset = 0, tries = 0;
  while (true) {
    const full = `${url}&limit=${limit}&offset=${offset}&filter=stockStore=${storeHref}`;
    try {
      const resp = ms.fetch(full, { method: 'GET' });
      const data = JSON.parse(resp.getContentText());
      const rows = data.rows || [];
      for (const r of rows) {
        fin.push({
          code: r && r.code ? String(r.code) : '',
          stock: r && r.stock ? Number(r.stock) : 0,
          inTransit: r && r.inTransit ? Number(r.inTransit) : 0
        });
      }
      if (rows.length < limit) break;
      offset += limit;
      tries = 0;
    } catch (e) {
      tries++;
      if (tries > 10) throw e;
      Utilities.sleep(tries * tries * 1000);
    }
  }
  return fin;
}

/* ============== Класс MoySklad (минимально нужное) ============== */
class MoySklad {
  constructor(token) {
    this.headers = { Authorization: `Bearer ${token}` };
  }
  fetch(url, opts) {
    const params = {
      headers: this.headers,
      muteHttpExceptions: true,
      method: (opts && opts.method) || 'GET'
    };
    let attempts = 0;
    while (true) {
      try {
        const resp = UrlFetchApp.fetch(url, params);
        const text = resp.getContentText();
        if (!text || text.trim().startsWith('<')) {
          throw new Error('MS returned non-JSON');
        }
        return resp;
      } catch (err) {
        attempts++;
        if (attempts > 10) throw err;
        Utilities.sleep(attempts * attempts * 1000);
      }
    }
  }
}

/* ============== Общие утилиты ============== */
function mustSheet(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Лист "${name}" не найден`);
  return sh;
}
function headerMap(row) {
  return row.reduce((m, v, i) => (v != null && v !== '' ? (m[String(v).trim()] = i, m) : m), {});
}
function mustHave(hdr, col, sheetName) {
  if (!(col in hdr)) throw new Error(`На листе "${sheetName}" отсутствует столбец "${col}"`);
}
function ensureCols(sheet, n) {
  const have = sheet.getMaxColumns();
  if (have < n) sheet.insertColumnsAfter(have, n - have);
}
function clearBlock(sheet, row, col, numRows, numCols) {
  const maxRows = sheet.getMaxRows();
  const r = Math.max(1, Math.min(row, maxRows));
  const nr = Math.max(0, Math.min(numRows, maxRows - r + 1));
  if (nr > 0) sheet.getRange(r, col, nr, numCols).clearContent();
}
function toNum(x) {
  const n = Number(x);
  return isFinite(n) ? n : NaN;
}
// Логическое приведение для доп. поля «Выключен»
function toBool(v) {
  if (typeof v === 'boolean') return v;
  const s = String(v).trim().toLowerCase();
  return ['1','true','да','yes','y','on','выкл','disabled'].includes(s);
}
