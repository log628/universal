/** =========================================================
 * getFIZICAL_OZ.gs (логгинг + сортировка заказов по дате ↓)
 *  - fiz1_OZ:  [OZ] Физ. оборот A:I ← пары из «[OZ] Артикулы»
 *  - fiz2_OZ:  [OZ] Физ. оборот K:S ← заказы 30д до конца вчера, сортировка по дате (L) DESC
 *  - fiz3_OZ:  [OZ] Физ. оборот U:Y ← остатки (>0) по /v1/analytics/stocks
 *  - fiz0_OZ:  оркестрация с подробным логгированием
 * ========================================================= */

/* ---------- Безопасные фолбэки, если REF не прогрузился ---------- */
const SHEET_FIZ_OZ   = (typeof REF !== 'undefined' && REF.SHEETS && REF.SHEETS.FIZ_OZ)   ? REF.SHEETS.FIZ_OZ   : '[OZ] Физ. оборот';
const SHEET_ARTS_OZ  = (typeof REF !== 'undefined' && REF.SHEETS && REF.SHEETS.ARTS_OZ)  ? REF.SHEETS.ARTS_OZ  : '[OZ] Артикулы';
const SHEET_PARAMS_OZ   = (typeof REF !== 'undefined' && REF.SHEETS && REF.SHEETS.PARAMS)   ? REF.SHEETS.PARAMS   : '⚙️ Параметры';

function _readCabinetColorMapAll_(platform) {
  try {
    return REF.readCabinetColorMap(platform || 'OZON'); // фильтр строго по D
  } catch(_) { return new Map(); }
}


/* =========================
 *        З А П У С К
 * ========================= */
function fiz0_OZ() {
  const ss = SpreadsheetApp.getActive();
  const T0 = Date.now();
  const log = (label, extra) => {
    const ms = String(Date.now() - T0).padStart(7,' ');
    console.log(`[${ms} ms] fiz0_OZ | ${label}${extra ? ' | ' + extra : ''}`);
  };

  ss.toast('Остатки → Заказы → Расчет', 'Запуск', 3);
  log('START');

  // 1) Остатки
  let t = Date.now();
  let res3 = {};
  try {
    res3 = fiz3_OZ() || {};
    log('fiz3_OZ OK', `rows=${res3.rows||0}, cabinets=${res3.cabs||0}, took=${Date.now()-t}ms`);
  } catch (e) {
    log('fiz3_OZ FAIL', e && e.message ? e.message : String(e));
  }

  // 2) Заказы
  t = Date.now();
  let res2 = {};
  try {
    res2 = fiz2_OZ() || {};
    log('fiz2_OZ OK', `rows=${res2.rows||0}, cabinets=${res2.cabs||0}, took=${Date.now()-t}ms`);
  } catch (e) {
    log('fiz2_OZ FAIL', e && e.message ? e.message : String(e));
  }

  // 3) Расчет
  t = Date.now();
  let res1 = {};
  try {
    res1 = fiz1_OZ() || {};
    log('fiz1_OZ OK', `rows=${res1.rows||0}, pairs=${res1.pairs||0}, took=${Date.now()-t}ms`);
  } catch (e) {
    log('fiz1_OZ FAIL', e && e.message ? e.message : String(e));
  }

  const took = Date.now() - T0;
  ss.toast('Обновлено', SHEET_FIZ_OZ, 3);
  log('END', `took=${took}ms | summary={UY:${res3.rows||0}, KS:${res2.rows||0}, AI:${res1.rows||0}}`);
}

/* =========================
 *          fiz1_OZ
 * ========================= */
function fiz1_OZ() {
  const ss = SpreadsheetApp.getActive();
  const T0 = Date.now();
  const log = (label, extra) => {
    const ms = String(Date.now() - T0).padStart(7,' ');
    console.log(`[${ms} ms] fiz1_OZ | ${label}${extra ? ' | ' + extra : ''}`);
  };

  const shDst  = ss.getSheetByName(SHEET_FIZ_OZ) || ss.insertSheet(SHEET_FIZ_OZ);
  const shArts = ss.getSheetByName(SHEET_ARTS_OZ);
  if (!shArts) throw new Error('Лист «' + SHEET_ARTS_OZ + '» не найден');

  _ensureFizHeader_OZ(shDst);
  log('HEADER ensured', `sheet=${SHEET_FIZ_OZ}`);

  // 1) пары
  const pick = _getSelectedCabinetsFromParamsFIZ();
  const filterActive = pick.selected.length && pick.selected.length < pick.total;
  log('PARAMS', `total=${pick.total}, selected=${pick.selected.length}, mode=${filterActive?'multi':'all/none'}`);

  const artsLast = shArts.getLastRow();
  const artsRows = artsLast > 1 ? shArts.getRange(2, 1, artsLast - 1, 2).getDisplayValues() : [];
  const keyList = [];
  const keySet  = new Set();
  for (const r of artsRows) {
    const cab = String(r[0] || '').trim();
    const art = String(r[1] || '').trim();
    if (!cab || !art) continue;
    if (filterActive && pick.selected.indexOf(cab) === -1) continue;
    const key = cab + '|' + art;
    if (!keySet.has(key)) { keySet.add(key); keyList.push({ cab, art }); }
  }
  // сортируем пары Кабинет→Артикул
  keyList.sort((a,b) => (a.cab.localeCompare(b.cab,'ru') || a.art.localeCompare(b.art,'ru')));
  log('PAIRS collected', `pairs=${keyList.length}`);

  // 2) Остатки из U:Y
  const fboSumByKey = new Map();
  const lastDst = shDst.getLastRow();
  const stocksBody = lastDst > 1 ? shDst.getRange(2, 21, lastDst - 1, 5).getValues() : []; // U:Y
  for (const r of stocksBody) {
    const cab = String(r[0] || '').trim();
    const art = String(r[1] || '').trim();
    const qty = _num0(r[2]);
    if (!cab || !art || qty <= 0) continue;
    fboSumByKey.set(cab + '|' + art, (fboSumByKey.get(cab + '|' + art) || 0) + qty);
  }
  log('STOCKS aggregated', `pairs=${fboSumByKey.size}`);

  // 3) Скорости (из K:S)
  const speedsByKey = new Map();
  const toEdge = _edgesToEndOfYesterday().to;
  const wins = [7,14,21,28];
  if (lastDst > 1) {
    const ordBody = shDst.getRange(2, 11, lastDst - 1, 9).getValues(); // K:S
    for (const r of ordBody) {
      const cab = String(r[0] || '').trim();                                     // K
      const dt  = r[1] instanceof Date ? r[1] : (r[1] ? new Date(r[1]) : null);  // L
      const art = String(r[3] || '').trim();                                     // N
      const qty = _num0(r[4]);                                                   // O
      const method = String(r[5] || '').trim().toUpperCase();                    // P
      if (!cab || !art || !dt || !isFinite(dt.getTime()) || qty <= 0) continue;

      const key = cab + '|' + art;
      if (!speedsByKey.has(key)) {
        speedsByKey.set(key, { w7:{sum:0,hasFbs:false}, w14:{sum:0,hasFbs:false}, w21:{sum:0,hasFbs:false}, w28:{sum:0,hasFbs:false} });
      }
      const rec = speedsByKey.get(key);
      for (const w of wins) {
        const wEdge = _windowEdges(w, toEdge);
        if (dt >= wEdge.since && dt <= wEdge.to) {
          const slot = w===7?'w7':w===14?'w14':w===21?'w21':'w28';
          rec[slot].sum += qty;
          if (method === 'FBS') rec[slot].hasFbs = true;
        }
      }
    }
  }
  log('SPEEDS computed', `pairs=${speedsByKey.size}`);

  // 4) Формирование и запись A:I
  const out = [];
  const fmts = [];
  const FMT_PLAIN = '0.00';
  const FMT_BRACK = '"["0.00"]"';

  for (const {cab, art} of keyList) {
    const key = cab + '|' + art;

    const fboLeftNum = _num0(fboSumByKey.get(key) || 0);
    const fboLeft = fboLeftNum === 0 ? '' : fboLeftNum;

    const rec = speedsByKey.get(key) || { w7:{sum:0,hasFbs:false}, w14:{sum:0,hasFbs:false}, w21:{sum:0,hasFbs:false}, w28:{sum:0,hasFbs:false} };
    const vals = [
      {v: rec.w7.sum/7,   hasFbs: rec.w7.hasFbs},
      {v: rec.w14.sum/14, hasFbs: rec.w14.hasFbs},
      {v: rec.w21.sum/21, hasFbs: rec.w21.hasFbs},
      {v: rec.w28.sum/28, hasFbs: rec.w28.hasFbs},
    ];

    const hVals = vals.map(o => (o.v || 0) === 0 ? '' : Number(o.v.toFixed(2)));
    const hFmts = vals.map(o => ((o.v || 0) === 0 ? FMT_PLAIN : (o.hasFbs ? FMT_BRACK : FMT_PLAIN)));

    let maxV = 0, maxHasFbs = false;
    for (const o of vals) if ((o.v||0) > maxV) { maxV = o.v; maxHasFbs = o.hasFbs; }
    const speedVal = (maxV === 0 ? '' : Number(maxV.toFixed(2)));
    const speedFmt = (maxV === 0 ? FMT_PLAIN : (maxHasFbs ? FMT_BRACK : FMT_PLAIN));

    // A:I = Каб | Арт | Остаток | В поставке | Скорость | Ск[7] | Ск[14] | Ск[21] | Ск[28]
    out.push([ cab, art, fboLeft, '', speedVal, hVals[0], hVals[1], hVals[2], hVals[3] ]);
    fmts.push(['@','@','0.########','@', speedFmt, hFmts[0], hFmts[1], hFmts[2], hFmts[3]]);
  }

  _clearBlock(shDst, 2, 1, 99999, 9); // A:I
  if (out.length) {
    shDst.getRange(2, 1, out.length, 9).setValues(out);
    shDst.getRange(2, 1, out.length, 9).setNumberFormats(fmts);
    shDst.autoResizeColumn(2);
    shDst.setColumnWidth(2, shDst.getColumnWidth(2) + 30);
    _applyFizDataStyles(shDst, 2, out.length, 9, 10);
    _applyCabinetRowFills_OZ(shDst, out, 1, 9);
  }

  SpreadsheetApp.flush();
  ss.toast('Расчет: записано строк ' + out.length, 'Готово', 4);
  log('END', `rows=${out.length}`);

  return { rows: out.length, pairs: keyList.length };
}

/* =========================
 *          fiz2_OZ
 * ========================= */
function fiz2_OZ() {
  const ss = SpreadsheetApp.getActive();
  const T0 = Date.now();
  const log = (label, extra) => {
    const ms = String(Date.now() - T0).padStart(7,' ');
    console.log(`[${ms} ms] fiz2_OZ | ${label}${extra ? ' | ' + extra : ''}`);
  };

  const sh = ss.getSheetByName(SHEET_FIZ_OZ) || ss.insertSheet(SHEET_FIZ_OZ);
  _ensureFizHeader_OZ(sh);
  log('HEADER ensured', `sheet=${SHEET_FIZ_OZ}`);

  const edges = _getLast30DaysEdgesExclToday(); // {sinceISO,toISO}
  const pick  = _getSelectedCabinetsFromParamsFIZ();
  let accs = {};
  try { accs = OZONAPI.getAccounts() || {}; } catch(_) { accs = {}; }
  log('PARAMS', `total=${pick.total}, selected=${pick.selected.length}, cabinetsToProcess=${pick.cabinetsToProcess.length}`);

  const whToClusterCode = new Map();
  const out = [];
  let cabsProcessed = 0;

  for (const cab of pick.cabinetsToProcess) {
    if (!accs[cab]) { log('SKIP cab (no keys)', cab); continue; }
    cabsProcessed++;
    const tCab = Date.now();
    const api = new OZONAPI(cab);
    let fbo = [], fbs = [];
    try { fbo = api.getOrdersFbo(30) || []; } catch(e){ log('WARN getOrdersFbo', `${cab}: ${e.message||e}`); }
    try { fbs = api.getOrdersFbs(30) || []; } catch(e){ log('WARN getOrdersFbs', `${cab}: ${e.message||e}`); }

    _collectOrdersMT(out, cab, fbo, 'FBO', edges, whToClusterCode, /*useCodes=*/true);
    _collectOrdersMT(out, cab, fbs, 'FBS', edges, whToClusterCode, /*useCodes=*/false);
    log('CAB done', `${cab}: rowsNow=${out.length}, took=${Date.now()-tCab}ms`);
  }

  _clearBlock(sh, 2, 11, 99999, 9); // K:S
  if (out.length) {
    const rng = sh.getRange(2, 11, out.length, 9);
    rng.setValues(out);

    sh.getRange(2, 11, out.length, 9)
      .setFontFamily('Roboto').setFontSize(8).setFontColor('#000000')
      .setFontWeight('normal').setHorizontalAlignment('left').setVerticalAlignment('middle');

    // сортировка по дате ↓ (L)
    rng.sort([{ column: 12, ascending: false }]); // L
    log('SORT applied', 'by L (date) DESC');
  }

  SpreadsheetApp.flush();
  ss.toast('Заказы: записано строк ' + out.length, 'Готово', 3);
  log('END', `rows=${out.length}, cabs=${cabsProcessed}`);

  return { rows: out.length, cabs: cabsProcessed };
}


/* =========================
 *          fiz3_OZ
 * ========================= */
function fiz3_OZ() {
  const ss = SpreadsheetApp.getActive();
  const T0 = Date.now();
  const log = (label, extra) => {
    const ms = String(Date.now() - T0).padStart(7,' ');
    console.log(`[${ms} ms] fiz3_OZ | ${label}${extra ? ' | ' + extra : ''}`);
  };

  const sh = ss.getSheetByName(SHEET_FIZ_OZ) || ss.insertSheet(SHEET_FIZ_OZ);
  const shArts = ss.getSheetByName(SHEET_ARTS_OZ);
  if (!shArts) throw new Error('Лист «' + SHEET_ARTS_OZ + '» не найден');

  _ensureFizHeader_OZ(sh);
  log('HEADER ensured', `sheet=${SHEET_FIZ_OZ}`);

  const pick = _getSelectedCabinetsFromParamsFIZ();
  log('PARAMS', `total=${pick.total}, selected=${pick.selected.length}, cabinetsToProcess=${pick.cabinetsToProcess.length}`);

  // --- offersByCab ---
  const offersByCab = new Map();
  {
    const last = shArts.getLastRow();
    const rows = last > 1 ? shArts.getRange(2, 1, last - 1, 2).getDisplayValues() : [];
    const filterActive = pick.selected.length && pick.selected.length < pick.total;
    for (const r of rows) {
      const cab = String(r[0] || '').trim();
      const offer = String(r[1] || '').trim();
      if (!cab || !offer) continue;
      if (filterActive && pick.selected.indexOf(cab) === -1) continue;
      if (!offersByCab.has(cab)) offersByCab.set(cab, new Set());
      offersByCab.get(cab).add(offer);
    }
  }
  log('OFFERS grouped', `cabs=${offersByCab.size}`);

  let accs = {};
  try { accs = OZONAPI.getAccounts() || {}; } catch(_) { accs = {}; }
  const out = [];
  let cabsProcessed = 0;

  for (const [cab, offerSet] of offersByCab) {
    if (!accs[cab]) { log('SKIP cab (no keys)', cab); continue; }
    cabsProcessed++;
    const tCab = Date.now();
    const api = new OZONAPI(cab);
    const offers = Array.from(offerSet);
    if (!offers.length) continue;

    let offerToSku = {};
    try { offerToSku = api.getSkusByOffers(offers) || {}; } catch(e) { offerToSku = {}; log('WARN getSkusByOffers', `${cab}: ${e.message||e}`); }
    const skus = Object.values(offerToSku).filter(Boolean);
    if (!skus.length) { log('NO SKUS', cab); continue; }

    let items = [];
    try { items = api.analyticsStocksBySkus(skus) || []; } catch(e) { items = []; log('WARN analyticsStocksBySkus', `${cab}: ${e.message||e}`); }

    const skuToOffer = {};
    for (const off in offerToSku) {
      const sku = offerToSku[off];
      if (sku != null) skuToOffer[String(sku)] = off;
    }

    for (const it of items) {
      const sku = it && (it.sku ?? it.sku_id ?? it.id);
      const offer = (sku != null ? skuToOffer[String(sku)] : (it.offer_id || it.offer)) || '';
      if (!offer) continue;

      const qty = _num0(it.available_stock_count || it.available || it.qty);
      if (qty <= 0) continue;

      const whName  = String(it.warehouse_name || it.warehouse || '').trim();
      const cluster = String(it.cluster_name || it.cluster || '').trim();

      out.push([cab, offer, qty, whName, cluster]);
    }
    log('CAB done', `${cab}: rowsNow=${out.length}, took=${Date.now()-tCab}ms`);
  }

  _clearBlock(sh, 2, 21, 99999, 5); // U:Y
  if (out.length) {
    sh.getRange(2, 21, out.length, 5).setValues(out);
    sh.getRange(2, 21, out.length, 5)
      .setFontFamily('Roboto').setFontSize(8).setFontColor('#000000')
      .setFontWeight('normal').setHorizontalAlignment('left').setVerticalAlignment('middle');
  }

  SpreadsheetApp.flush();
  ss.toast('Остатки: записано строк ' + out.length, 'Готово', 3);
  log('END', `rows=${out.length}, cabs=${cabsProcessed}`);

  return { rows: out.length, cabs: cabsProcessed };
}


/* =========================
 *       Х Е Л П Е Р Ы
 * ========================= */

function _findHeaderIndex(headerRowValues, names) {
  const norm = (s) => String(s||'').trim().toLowerCase();
  const hdr = headerRowValues.map(norm);
  const candidates = names.map(norm);
  for (let i = 0; i < hdr.length; i++) {
    if (candidates.indexOf(hdr[i]) !== -1) return i;
  }
  return -1;
}

function _edgesToEndOfYesterday() {
  const now = new Date();
  const to = new Date(now);
  to.setDate(to.getDate() - 1);
  to.setHours(23,59,59,999);
  return { to };
}

function _windowEdges(days, to) {
  const since = new Date(to);
  since.setDate(since.getDate() - (days - 1));
  since.setHours(0,0,0,0);
  return { since, to };
}

function _num0(v) { return isFinite(+v) ? (+v) : 0; }

function _getLast30DaysEdgesExclToday() {
  const now = new Date();
  const to = new Date(now);
  to.setDate(to.getDate() - 1);
  to.setHours(23, 59, 59, 999);
  const since = new Date(to);
  since.setDate(since.getDate() - 29);
  since.setHours(0, 0, 0, 0);
  return { sinceISO: since.toISOString(), toISO: to.toISOString() };
}

function _clearBlock(sh, r, c, h, w) {
  sh.getRange(r, c, h, w).clear({ contentsOnly: true });
}

function _normWh(s) {
  return String(s || '').trim().replace(/\s+/g,' ').toLowerCase();
}

/* Хедеры: везде шрифт = 10; в A1 — "[ OZ ] Кабинет"
   Блоки:
   - A:I  — сводка (без «Склад»/«В пути»)
   - K:S  — заказы (с «ID отправления» после даты)
   - U:Y  — остатки */
function _ensureFizHeader_OZ(sh) {
  // --- A:I ---
  const headersAI = ['[ OZ ] Кабинет','Артикул','Остаток','В поставке','Скорость','Ск [ 7 ]','Ск [ 14 ]','Ск [ 21 ]','Ск [ 28 ]'];
  sh.getRange(1, 1, 1, headersAI.length).setValues([headersAI]);
  const rngAI = sh.getRange(1, 1, 1, headersAI.length);
  rngAI
    .setFontFamily('Roboto').setFontSize(10).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, 2).setBackground('#434343'); // A:B
  sh.getRange(1, 3, 1, 2).setBackground('#38761d'); // C:D
  sh.getRange(1, 5, 1, 1).setBackground('#1155cc'); // E
  sh.getRange(1, 6, 1, 4).setBackground('#6d9eeb'); // F:I
  for (const c of [1,3,4,5]) sh.setColumnWidth(c, 100);
  for (const c of [6,7,8,9]) sh.setColumnWidth(c, 65);
  sh.autoResizeColumn(2);
  sh.setColumnWidth(2, sh.getColumnWidth(2) + 30);

  // тег "[ OZ ]" — жёлтым
  try {
    const a1Text = headersAI[0];
    const a1 = sh.getRange('A1');
    const tagEnd = a1Text.indexOf(']') + 1;
    const builder = SpreadsheetApp.newRichTextValue().setText(a1Text);
    if (tagEnd > 0) {
      builder.setTextStyle(0, tagEnd, SpreadsheetApp.newTextStyle().setForegroundColor('#ffff00').setBold(true).build());
      builder.setTextStyle(tagEnd, a1Text.length, SpreadsheetApp.newTextStyle().setForegroundColor('#ffffff').setBold(true).build());
    }
    a1.setRichTextValue(builder.build());
  } catch (_) {}

  // --- K:S ---
  const headersKS = ['Кабинет','Дата','ID отправления','Артикул','Количество','Метод','Склад отгрузки','Кластер отгрузки','Кластер доставки'];
  sh.getRange(1, 11, 1, headersKS.length).setValues([headersKS]);
  sh.getRange(1, 11, 1, headersKS.length)
    .setFontFamily('Roboto').setFontSize(10).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBackground('#434343');

  // --- U:Y ---
  const headersUY = ['Кабинет','Артикул','Количество','Склад хранения','Кластер хранения'];
  sh.getRange(1, 21, 1, headersUY.length).setValues([headersUY]);
  sh.getRange(1, 21, 1, headersUY.length)
    .setFontFamily('Roboto').setFontSize(10).setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBackground('#434343');

  sh.setFrozenRows(1);
}


/* Построчные заливки по Кабинету (с фолбэком, если REF нет) */
/* Построчные заливки по Кабинету (с фильтром по площадке) */
function _applyCabinetRowFills_OZ(sh, rows, startCol, numCols) {
  try {
    const colorMap = _readCabinetColorMapAll_('OZON'); // ← ключевая правка
    if (!rows || !rows.length) return;

    const fills = [];
    for (let i = 0; i < rows.length; i++) {
      const cab = String(rows[i][0] || '').trim();
      const color = (colorMap && colorMap.get(cab)) || '#ffffff';

      // DEBUG: покажем, кого не нашли в карте цветов
      if (!colorMap.has(cab)) console.log('NO COLOR FOR:', JSON.stringify(cab));

      const line = new Array(numCols);
      for (let c = 0; c < numCols; c++) line[c] = color;
      fills.push(line);
    }
    sh.getRange(2, startCol, rows.length, numCols).setBackgrounds(fills);
  } catch (e) {}
}



/* Универсальные стили данных для блоков (по умолчанию шрифт 8) */
function _applyFizDataStyles(sh, startRow, numRows, numCols, fontSize /* =8 */) {
  if (numRows <= 0) return;
  const size = (fontSize && +fontSize) ? +fontSize : 8;
  sh.getRange(startRow, 1, numRows, numCols)
    .setFontFamily('Roboto').setFontSize(size).setFontWeight('normal').setFontColor('#000000')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
}

/** Собираем заказы в K:S (фильтр отмен, коды кластеров, + posting id) */
function _collectOrdersMT(outRows, cab, postings, method, edges, whToClusterCodeMap, useCodes) {
  if (!Array.isArray(postings)) return;
  const since = new Date(edges.sinceISO);
  const to    = new Date(edges.toISO);

  for (const p of postings) {
    const dtStr = p.in_process_at || p.shipment_date || p.shipped_at || p.created_at;
    if (!dtStr) continue;
    const dt = new Date(dtStr);
    if (!(dt >= since && dt <= to)) continue;

    const statusRaw = String(p.status || p.state || '').trim().toLowerCase();
    if (statusRaw === 'cancelled' || statusRaw === 'canceled' || /^cancel/.test(statusRaw)) {
      continue;
    }

    const postingId = String(p.posting_number || p.postingNumber || p.id || '').trim();

    const analytics = p.analytics_data || {};
    const whName = analytics.warehouse_name || analytics.warehouse || '';

    let codeFrom = '', codeTo = '';
    if (useCodes) {
      const fd = p.financial_data || {};
      codeFrom = String(fd.cluster_from || '').trim();
      codeTo   = String(fd.cluster_to || '').trim();
      if (whName && codeFrom) {
        whToClusterCodeMap.set(_normWh(whName), codeFrom);
      }
    }

    const pushRow = (offer, qty) => {
      if (!offer || !qty) return;
      const sCode = useCodes ? codeFrom : (whToClusterCodeMap.get(_normWh(whName)) || '');
      const dCode = useCodes ? codeTo   : '';
      // K:S = Каб, Дата, ID отправления, Артикул, Кол-во, Метод, Склад отгр., Кл. отгр., Кл. дост.
      outRows.push([ cab, new Date(dt), postingId, offer, qty, method, whName || '', sCode, dCode ]);
    };

    const prods = Array.isArray(p.products) ? p.products : [];
    if (prods.length) {
      for (const pr of prods) {
        const offer = pr.offer_id || pr.offer || '';
        const qty   = _num0(pr.quantity || pr.qty);
        pushRow(offer, qty);
      }
      continue;
    }

    const rps = Array.isArray(p.resulting_products) ? p.resulting_products : [];
    for (const rp of rps) {
      const offer = rp.offer_id || rp.offer || '';
      const qty   = _num0(rp.quantity);
      pushRow(offer, qty);
    }
  }
}

/* Читаем выбранные кабинеты OZ из «⚙️ Параметры» (и валидируем по OZONAPI.getAccounts) */
function _getSelectedCabinetsFromParamsFIZ() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_PARAMS_OZ);

  let accounts = {};
  try { accounts = OZONAPI.getAccounts() || {}; } catch (_) { accounts = {}; }
  const validNames = new Set(Object.keys(accounts));

  const res = { selected: [], total: 0, cabinetsToProcess: [] };

  if (!sh) {
    const arts = ss.getSheetByName(SHEET_ARTS_OZ);
    if (arts && arts.getLastRow() > 1) {
      const names = arts.getRange(2,1,arts.getLastRow()-1,1).getDisplayValues()
        .map(r => String(r[0]||'').trim())
        .filter(n => n && validNames.has(n));
      const uniq = Array.from(new Set(names));
      res.selected = uniq.slice();
      res.total = uniq.length;
      res.cabinetsToProcess = uniq.slice();
    }
    return res;
  }

  const last = sh.getLastRow();
  if (last <= 1) return res;

  const rows = sh.getRange(2, 1, last - 1, 8).getDisplayValues();

  const allOzon = [];
  for (let i = 0; i < rows.length; i++) {
    const name  = String(rows[i][0] || '').trim();
    const platf = String(rows[i][3] || '').trim().toUpperCase();
    const on    = String(rows[i][7] || '').trim().toUpperCase();
    const isOn  = (on === 'TRUE' || on === 'ИСТИНА' || on === 'ДА');

    if (name && platf === 'OZON' && validNames.has(name)) {
      allOzon.push(name);
      if (isOn) res.selected.push(name);
    }
  }

  const uniqAll = Array.from(new Set(allOzon));
  const uniqSel = Array.from(new Set(res.selected));
  res.total     = uniqAll.length;
  res.selected  = uniqSel;

  if (!uniqSel.length || uniqSel.length === uniqAll.length) {
    res.cabinetsToProcess = uniqAll;
  } else {
    res.cabinetsToProcess = uniqSel;
  }

  return res;
}
