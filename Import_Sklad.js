/** ===============================
 *  Import_Sklad.gs — автономный сборщик
 *  ===============================
 * Делает:
 *  1) Читает токен МС из 🍔 СС!AF2
 *  2) Тянет ассортимент из МойСклад (Код/Производитель/Модель/Вес/Выключен)
 *  3) Выгружает склады в 🍔 СС!AH:AK (Склад/Код/Доступно/Ожидание)
 *  4) Собирает «Приёмки» (внешний файл), берёт Очередность=1
 *  5) Считает комплекты (Q:R:S):
 *      - СС в валюте (юань) = Σ(СС(состав, юань) * кол-во)
 *      - Наличие = min_i floor(наличие(part_i)/qty_i)   ← ТОЛЬКО ДЛЯ «Наличие»
 *      - «В пути», «В поставке OZ», «В поставке WB» — только прямые (без комплектов)
 *  6) Считает «СС+Упак+Дост»:
 *      - СС (руб) + Упаковка (руб, из AA) + Доставка(вес_кг * тариф$ * курс$ *1.1)
 *        (вес берём в граммах → кг; тариф «доставка» и «доллар» из N:O)
 *  7) Пишет итог в 🍔 СС!A:L
 *       A:Товар B:Производитель C:Модель D:СС в валюте E:Валюта
 *       F:СС+Упак+Дост G:Наличие H:В пути
 *       I:Остаток OZ   J:Остаток WB
 *       K:Сумма СС в руб (= (G+H+I+J)×D×курс(E), без дробей в отображении)
 *       L:Не закупается («да», если «Выключен» в асс-те МС)
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
    // === 0) Токен МС (AF2)
    const token = String(shCC.getRange('AF2').getValue() || '').trim();
    if (!token) throw new Error('Пустой токен МС в 🍔 СС!AF2');

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

    // === 2) Выгрузка складов → 🍔 СС!AH:AK
    const stockRows = exportStocksToCC_(ms, shCC); // [ [store, code, available, inTransit], ... ]
    let stockAgg  = aggregateStocks_(stockRows);    // code -> {availMain, transitMain, vpostOZ, vpostWB}

    // === 3) Курсы/тарифы N:O (юань/доллар/доставка/симка)
    const lastRowCC = shCC.getLastRow();
    const rates = {};
    if (lastRowCC >= 1) {
      const no = shCC.getRange(1, 14, lastRowCC, 2).getValues(); // N(14):O(15)
      for (const [name, val] of no) {
        const k = String(name || '').trim().toLowerCase();
        const v = toNum(val);
        if (!k || !isFinite(v)) continue;
        if (k === 'юань' || k === 'доллар' || k === 'доставка' || k.indexOf('симка')>-1) rates[k] = v;
      }
    }

    // === 4) «Упаковка» U:AA
    const packByCode = {};
    if (lastRowCC >= 2) {
      const pack = shCC.getRange(2, 21, lastRowCC - 1, 7).getValues(); // U..AA
      for (const row of pack) {
        const code = String(row[0] || '').trim(); // U
        if (!code) continue;
        const add = toNum(row[6]);                // AA
        packByCode[code] = isFinite(add) ? add : 0;
      }
    }

    // === 5) Комплекты Q:R:S
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
        const theCur = info ? (info.curr || '') : '';
        if (cost === '' || theCur !== 'юань') { ok = false; break; }
        sumYuan += Number(cost) * Number(qty || 0);
      }
      if (ok) priceByCode[kit] = { costDoc: sumYuan, curr: 'юань' };
    }

    // === 9) Комплектное «Наличие» (ТОЛЬКО availMain!)
    stockAgg = applyKitAvailOnly_(kits, stockAgg);

    // === 9.5) Остатки OZ/WB из листов физ.оборота
    // Формат: A=Кабинет, B=Артикул, C=Остаток; Артикул → Товар: cut 3 + remove "_catX"
    const mapOZ = readFizStocksMap_(REF.sheetName('FIZ_OZ'));
    const mapWB = readFizStocksMap_(REF.sheetName('FIZ_WB'));

    // === 10) Сборка итоговых строк A:L
    const HEADER = [
      'Товар','Производитель','Модель','СС в валюте','Валюта',
      'СС+Упак+Дост','Наличие','В пути',
      'Остаток OZ','Остаток WB',
      'Сумма СС в руб','Не закупается'
    ];

    // курс по валюте
    const getRate = (curr) => {
      const c = String(curr || '').toLowerCase();
      if (!c) return NaN;
      if (c === 'рубль' || c === 'rub' || c === 'rur') return 1;
      if (c === 'юань') return toNum(rates['юань']);
      if (c === 'доллар') return toNum(rates['доллар']);
      return NaN;
    };

    const codes = Array.from(codeSet).filter(Boolean)
      .sort((a,b)=>String(a).localeCompare(String(b)));

    const out = [];
    let sumG = 0, sumH = 0; // итоги G/H
    let sumI = 0, sumJ = 0; // итоги I/J

    for (const code of codes) {
      const p = prodByCode[code] || {manufacturer:'', model:'', weightRaw:'', disabled:false};
      const r = priceByCode[code] || {costDoc:'', curr:''};
      const s = stockAgg[code]    || {availMain:0, transitMain:0, vpostOZ:0, vpostWB:0};

      // F: СС в рублях (база: D×курс + упаковка + доставка)
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

      // G/H: нули -> пусто
      const gVal = Number(s.availMain) || 0;
      const hVal = Number(s.transitMain) || 0;
      const gDisp = gVal === 0 ? '' : gVal;
      const hDisp = hVal === 0 ? '' : hVal;
      sumG += gVal;
      sumH += hVal;

      // I/J: остатки из карт (0 -> пусто)
      const iVal = Number(mapOZ.get(code) || 0);
      const jVal = Number(mapWB.get(code) || 0);
      const iDisp = iVal === 0 ? '' : iVal;
      const jDisp = jVal === 0 ? '' : jVal;
      sumI += iVal;
      sumJ += jVal;

      // K: «Сумма СС в руб» = (G+H+I+J) × D × курс(E), 0 → пусто
      const units   = gVal + hVal + iVal + jVal;
      const unitCost = toNum(r.costDoc);
      const rate    = getRate(r.curr);
      let kVal = (isFinite(unitCost) && isFinite(rate)) ? ( (units || 0) * Number(unitCost) * Number(rate) ) : NaN;
      const kDisp = (!isFinite(kVal) || Math.round(kVal) === 0) ? '' : Math.round(kVal);

      const notPurchasingText = p.disabled ? 'да' : '';

      out.push([
        code, p.manufacturer, p.model,
        r.costDoc === '' ? '' : Number(r.costDoc),
        curr || '',
Math.round(totalRub),

        gDisp, hDisp,
        iDisp, jDisp,
        kDisp,
        notPurchasingText
      ]);
    }

    // === 10.5) Сортировка результата
    out.sort((r1, r2) => {
      const l1 = String(r1[11] || '').trim() === '' ? 0 : 1; // L (index 11)
      const l2 = String(r2[11] || '').trim() === '' ? 0 : 1;
      if (l1 !== l2) return l1 - l2;
      return String(r1[0] || '').localeCompare(String(r2[0] || '')); // A: Товар
    });

    // === 11) Запись в 🍔 СС!A:L
    clearBlock(shCC, 1, 1, shCC.getMaxRows(), 12);
    ensureCols(shCC, 12);
    shCC.getRange(1,1,1,12).setValues([HEADER]);
    if (out.length) shCC.getRange(2,1,out.length,12).setValues(out);

    shCC.setFrozenRows(1);

    // === 11.x) ВЫРАВНИВАНИЯ
    // 1) Заголовки A1:L1 — по центру
    shCC.getRange(1, 1, 1, 12).setHorizontalAlignment('center');

    if (out.length) {
      // 2) Все данные (A2:L...) — по левому краю
      shCC.getRange(2, 1, out.length, 12).setHorizontalAlignment('left');

      // === 11.y) ЧИСЛОВЫЕ И ТЕКСТОВЫЕ ФОРМАТЫ С ОТСТУПОМ «2 ПРОБЕЛА» (числа остаются числами)
      // Текстовые столбцы: A,B,C,E,L -> "  "@
      shCC.getRange(2,1, out.length,1).setNumberFormat('"  "@'); // A
      shCC.getRange(2,2, out.length,1).setNumberFormat('"  "@'); // B
      shCC.getRange(2,3, out.length,1).setNumberFormat('"  "@'); // C
      shCC.getRange(2,5, out.length,1).setNumberFormat('"  "@'); // E
      shCC.getRange(2,12,out.length,1).setNumberFormat('"  "@'); // L

      // Денежные с десятичными: D,F -> "  "#,##0.##
      shCC.getRange(2,4, out.length,1).setNumberFormat('"  "#,##0.##'); // D
shCC.getRange(2,6, out.length,1).setNumberFormat('"  "#,##0'); // F — без десятичных


      // Целые: G,H,I,J,K -> "  "#,##0
      shCC.getRange(2,7, out.length,1).setNumberFormat('"  "#,##0'); // G
      shCC.getRange(2,8, out.length,1).setNumberFormat('"  "#,##0'); // H
      shCC.getRange(2,9, out.length,1).setNumberFormat('"  "#,##0'); // I
      shCC.getRange(2,10,out.length,1).setNumberFormat('"  "#,##0'); // J
      shCC.getRange(2,11,out.length,1).setNumberFormat('"  "#,##0'); // K

      // Заливки данных + шрифты
      shCC.getRange(2, 7, out.length, 2).setBackground('#fff6ed'); // G:H
      shCC.getRange(2, 9, out.length, 1).setBackground('#e0ebff').setFontColor('#000000').setFontWeight('bold'); // I (OZ)
      shCC.getRange(2,10, out.length, 1).setBackground('#e9e6f1').setFontColor('#000000').setFontWeight('bold'); // J (WB)
      shCC.getRange(2,11, out.length, 1).setBackground('#effcf3'); // K (сумма)
    }

    // === 11.1) Заголовок K1 — двухстрочный + стили (сумма без дробной части)
    const kCol = 11; // K
    const kHeaderCell = shCC.getRange(1, kCol);
    const totalKRub = Math.round(out.reduce((acc, r) => acc + (isFinite(Number(r[10])) ? Number(r[10]) : 0), 0));
    const titleK = 'Сумма СС в руб';
    const sumKStr = fmtMoneyRU_(totalKRub);
    const fullK = titleK + '\n' + sumKStr;

    const whiteBold  = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#ffffff').build();
    const yellowBold = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#ffff00').build();
    const blackBold  = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#000000').build();

    const kRTV = SpreadsheetApp.newRichTextValue()
      .setText(fullK)
      .setTextStyle(0, titleK.length, whiteBold)
      .setTextStyle(titleK.length + 1, fullK.length, yellowBold)
      .build();

    kHeaderCell.setRichTextValue(kRTV)
      .setBackground('#34a853')
      .setWrap(true)
      .setVerticalAlignment('middle');

    // === 11.2) Двухстрочные заголовки для G1/H1 (итоги)
    const gTitle = 'Наличие';
    const hTitle = 'В пути';
    const gTotalStr = fmtIntRU_(sumG);
    const hTotalStr = fmtIntRU_(sumH);

    const brownBold = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#b45f06').build();

    // G1
    const gCell = shCC.getRange(1, 7);
    const gText = gTitle + '\n' + gTotalStr;
    const gRTV = SpreadsheetApp.newRichTextValue()
      .setText(gText)
      .setTextStyle(0, gTitle.length, blackBold)
      .setTextStyle(gTitle.length + 1, gText.length, brownBold)
      .build();
    gCell.setRichTextValue(gRTV).setWrap(true).setVerticalAlignment('middle');

    // H1
    const hCell = shCC.getRange(1, 8);
    const hText = hTitle + '\n' + hTotalStr;
    const hRTV = SpreadsheetApp.newRichTextValue()
      .setText(hText)
      .setTextStyle(0, hTitle.length, blackBold)
      .setTextStyle(hTitle.length + 1, hText.length, brownBold)
      .build();
    hCell.setRichTextValue(hRTV).setWrap(true).setVerticalAlignment('middle');

    // === 11.3) Двухстрочные заголовки для I1/J1 (Остатки) + фоны
    // I1 — Остаток OZ
    const iTitle = 'Остаток OZ';
    const iTotalStr = fmtIntRU_(sumI);
    const iText = iTitle + '\n' + iTotalStr;
    const blueBold = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#0000ff').build();
    const iCell = shCC.getRange(1, 9);
    const iRTV = SpreadsheetApp.newRichTextValue()
      .setText(iText)
      .setTextStyle(0, iTitle.length, blackBold)
      .setTextStyle(iTitle.length + 1, iText.length, blueBold)
      .build();
    iCell.setRichTextValue(iRTV)
      .setBackground('#a4c2f4')
      .setWrap(true)
      .setVerticalAlignment('middle');

    // J1 — Остаток WB
    const jTitle = 'Остаток WB';
    const jTotalStr = fmtIntRU_(sumJ);
    const jText = jTitle + '\n' + jTotalStr;
    const purpleBold = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#9900ff').build();
    const jCell = shCC.getRange(1, 10);
    const jRTV = SpreadsheetApp.newRichTextValue()
      .setText(jText)
      .setTextStyle(0, jTitle.length, blackBold)
      .setTextStyle(jTitle.length + 1, jText.length, purpleBold)
      .build();
    jCell.setRichTextValue(jRTV)
      .setBackground('#d9d2e9')
      .setWrap(true)
      .setVerticalAlignment('middle');

    // === 11.4) Грани:
    const usedRows = Math.max(1, 1 + out.length);

    // Между C:D (правый бордер у C)
    shCC.getRange(1, 3, usedRows, 1)
      .setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Между F:G (правый бордер у F)
    shCC.getRange(1, 6, usedRows, 1)
      .setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Колонка K — лево/право средняя чёрная по всей высоте
    shCC.getRange(1, 11, usedRows, 1)
      .setBorder(null, true, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Нижняя граница у последней строки для A:J
    shCC.getRange(usedRows, 1, 1, 10)
      .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Нижняя граница у последней ячейки K
    shCC.getRange(usedRows, 11, 1, 1)
      .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

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

// --- МС: ассортимент
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
          else if (a.name === 'Вес' && 'value' in a)      weightRaw = a.value; // граммы
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

// --- Выгрузка складов в 🍔 СС!AH:AK, и вернуть «сырые» строки
function exportStocksToCC_(ms, shCC) {
  // Заголовки AH:AK
  shCC.getRange('AH1:AK1').setValues([['Склад','Код','Доступно','Ожидание']]);

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
  if (maxRows > 1) shCC.getRange(2, 34, maxRows - 1, 4).clearContent(); // AH=34
  if (rows.length) shCC.getRange(2, 34, rows.length, 4).setValues(rows);

  return rows; // [ [store, code, available, inTransit], ... ]
}

// --- Агрегаты по складам из сырья
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

// --- Комплекты Q:R:S → {kit: [{part, qty}, ...]}
function readKits_(shCC) {
  const last = shCC.getLastRow();
  const res = {};
  if (last < 2) return res;
  const qrs = shCC.getRange(2, 17, last - 1, 3).getValues(); // Q(17):S(19)
  for (const [kitRaw, partRaw, qtyRaw] of qrs) {
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

// --- Прочитать карту остатков из листа физ.оборота: A=Кабинет, B=Артикул, C=Остаток
// Артикул -> Товар: убираем первые 3 символа, снимаем суффикс "_catX" (X — одна цифра)
function readFizStocksMap_(sheetName) {
  const map = new Map();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return map;

  const last = sh.getLastRow();
  if (last < 2) return map;

  const vals = sh.getRange(2, 1, last - 1, 3).getDisplayValues(); // A..C
  for (let i = 0; i < vals.length; i++) {
    const art = String(vals[i][1] || '').trim();        // B: Артикул
    if (!art) continue;

    // Артикул -> Товар
    let tovar = art.length >= 3 ? art.substring(3) : '';
    tovar = tovar.replace(/_cat\d$/i, '');              // снимаем суффикс "_catX" (одна цифра)
    tovar = tovar.trim();
    if (!tovar) continue;

    // Остаток (C)
    const stockRaw = vals[i][2];
    const n = (typeof REF !== 'undefined' && typeof REF.toNumber === 'function')
      ? REF.toNumber(stockRaw)
      : (isFinite(Number(String(stockRaw).replace(',', '.'))) ? Number(String(stockRaw).replace(',', '.')) : 0);

    map.set(tovar, isFinite(n) ? Number(n) : 0);
  }
  return map;
}

// --- МС helpers
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
  constructor(token) { this.headers = { Authorization: `Bearer ${token}` }; }
  fetch(url, opts) {
    const params = { headers: this.headers, muteHttpExceptions: true, method: (opts && opts.method) || 'GET' };
    let attempts = 0;
    while (true) {
      try {
        const resp = UrlFetchApp.fetch(url, params);
        const text = resp.getContentText();
        if (!text || text.trim().startsWith('<')) { throw new Error('MS returned non-JSON'); }
        return resp;
      } catch (err) {
        attempts++; if (attempts > 10) throw err;
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
function toBool(v) {
  if (typeof v === 'boolean') return v;
  const s = String(v).trim().toLowerCase();
  return ['1','true','да','yes','y','on','выкл','disabled'].includes(s);
}
function fmtMoneyRU_(n) {
  try {
    return new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 0, maximumFractionDigits: 0 })
      .format(Math.round(n || 0));
  } catch (e) {
    const v = Math.round(n || 0);
    return String(v).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  }
}
function fmtIntRU_(n) {
  try {
    return new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 }).format(Math.round(n || 0));
  } catch (e) {
    const v = Math.round(n || 0);
    return String(v).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  }
}


/**
 * Лёгкий запуск импорта склада — обновляет ТОЛЬКО:
 *  - сырьё по складам в 🍔 СС!AH:AK
 *  - колонку G (Наличие) в основном списке A:L
 *  - вторую строку заголовка G1 (итог)
 *
 * Ничего больше не трогает: A–F, H–L, стили и сортировку не меняет.
 */
function Import_Sklad_GHOnly() {} // оставлен специально, чтобы не было коллизий имён
function Import_Sklad_GHOnly() {
  const ss   = SpreadsheetApp.getActive();
  const shCC = mustSheet(ss, '🍔 СС');

  const T0 = Date.now();
  const t = (label) => `[+${String(Date.now()-T0).padStart(6,' ')} ms] ${label}`;
  console.log(t('START Import_Sklad_GOnly'));
  ss.toast('Импорт МойСклад → обновление G (Наличие)', 'Быстрый импорт', 3);

  try {
    // === 0) Токен МС (AF2)
    const token = String(shCC.getRange('AF2').getValue() || '').trim();
    if (!token) throw new Error('Пустой токен МС в 🍔 СС!AF2');

    // === 1) МойСклад клиент
    const ms = new MoySklad(token);

    // === 2) Выгрузка складов → 🍔 СС!AH:AK (сырьё)
    const stockRows = exportStocksToCC_(ms, shCC); // [ [store, code, available, inTransit], ... ]

    // === 3) Агрегация по кодам
    let stockAgg = aggregateStocks_(stockRows); // code -> {availMain, transitMain, vpostOZ, vpostWB}

    // === 4) Комплекты (ТОЛЬКО к G / availMain)
    const kits = readKits_(shCC);
    stockAgg = applyKitAvailOnly_(kits, stockAgg);

    // === 5) Обновление ТОЛЬКО G по существующему списку (A:L)
    const last = shCC.getLastRow();
    if (last >= 2) {
      const codes = shCC.getRange(2, 1, last-1, 1).getValues(); // A
      const outG = new Array(codes.length);
      let sumG = 0;

      for (let i=0; i<codes.length; i++){
        const code = String(codes[i][0] || '').trim();
        const rec  = code ? stockAgg[code] : null;
        const gVal = rec ? Number(rec.availMain)||0 : 0;

        // отображение: нули → пусто
        outG[i] = [ gVal === 0 ? '' : gVal ];

        // для итога — сумму считаем по числам
        sumG += gVal;
      }

      // Запись G
      shCC.getRange(2, 7, outG.length, 1).setValues(outG);
      // Локальная заливка как в полном импорте (только колонка G)
      shCC.getRange(2, 7, outG.length, 1).setBackground('#fff6ed');

      // === 6) Обновить вторую строку заголовка G1 (итог)
      const blackBold = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#000000').build();
      const brownBold = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#b45f06').build();

      const gTitle = 'Наличие';
      const gText  = gTitle + '\n' + fmtIntRU_(sumG);
      const gRTV = SpreadsheetApp.newRichTextValue()
        .setText(gText)
        .setTextStyle(0, gTitle.length, blackBold)
        .setTextStyle(gTitle.length + 1, gText.length, brownBold)
        .build();
      shCC.getRange(1, 7).setRichTextValue(gRTV).setWrap(true).setVerticalAlignment('middle');
    }

    // === 7) Лог и финал
    safeLogRun_MS_(['МойСклад (G)']);
    ss.toast('Склад + СС (G): обновлено', 'Готово', 3);
    console.log(t('END Import_Sklad_GOnly'));
  } catch (e) {
    safeLogRun_MS_([]);
    ss.toast('Склад + СС (G): ошибка, см. журнал', 6);
    console.error(t(`ERROR Import_Sklad_GOnly: ${e && e.stack || e}`));
    throw e;
  }
}
