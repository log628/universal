/* =========================
 *        З А П У С К
 * ========================= */
function fiz0_WB() {
  const ss = SpreadsheetApp.getActive();
  const T0 = Date.now();
  const t = (label) => `[+${String(Date.now()-T0).padStart(6,' ')} ms] ${label}`;

  console.log(t('START fiz0_WB'));
  ss.toast('Остатки → Заказы → Расчет', SHEET_FIZ_WB, 3);

  let r3 = {}, r2 = {};
  try {
    const t1 = Date.now(); r3 = fiz3_WB() || {}; console.log(t(`fiz3_WB done in ${Date.now()-t1} ms; okCabs=${(r3.okCabs||[]).length}`));
    const t2 = Date.now(); r2 = fiz2_WB() || {}; console.log(t(`fiz2_WB done in ${Date.now()-t2} ms; okCabs=${(r2.okCabs||[]).length}`));
    const t3 = Date.now();      fiz1_WB();        console.log(t(`fiz1_WB done in ${Date.now()-t3} ms`));

    // --- ЛОГ В ⚙️ Параметры: только кабинеты, где были строки (заказы/остатки)
    try {
      const seen = new Set();
      const ok = [];
      for (const arr of [r2 && r2.okCabs || [], r3 && r3.okCabs || []]) {
        for (const cab of arr) { if (!seen.has(cab)) { seen.add(cab); ok.push(cab); } }
      }
      REF.logRun('Физ.оборот WB', ok, 'WILDBERRIES');
    } catch(_) {}

    ss.toast('Обновлено', SHEET_FIZ_WB, 3);
  } catch (e) {
    console.error(t(`ERROR fiz0_WB: ${e && e.stack || e}`));
    ss.toast('Ошибка. См. журнал.', SHEET_FIZ_WB, 5);
    throw e;
  } finally {
    console.log(t('END fiz0_WB'));
  }
}

/* =========================
 *          fiz1_WB
 * ========================= */
function fiz1_WB() {
  const ss = SpreadsheetApp.getActive();
  const shDst  = ss.getSheetByName(SHEET_FIZ_WB)  || ss.insertSheet(SHEET_FIZ_WB);
  const shArts = ss.getSheetByName(SHEET_ARTS_WB) || ss.insertSheet(SHEET_ARTS_WB);
  _ensureFizHeader_WB(shDst);

  // пары Кабинет/Артикул (с учётом чекбоксов в ⚙️ Параметры)
  const pick = _getSelectedCabinetsFromParamsFIZ_WB();
  const filterActive = pick.selected.length && pick.selected.length < pick.total;

  const artsLast = shArts.getLastRow();
  const artsRows = artsLast > 1 ? shArts.getRange(2, 1, artsLast - 1, 2).getDisplayValues() : [];
  const keyList = [];
  const keySet  = new Set();

  for (const r of artsRows) {
    const cab = String(r[0] || '').replace(/\u00A0/g, ' ').trim();
    const art = String(r[1] || '').trim();
    if (!cab || !art) continue;
    if (filterActive && pick.selected.indexOf(cab) === -1) continue;
    const key = cab + '|' + art;
    if (!keySet.has(key)) { keySet.add(key); keyList.push({ cab, art }); }
  }

  // Сортировка по Кабинет → Артикул
  keyList.sort((a, b) => {
    const byCab = a.cab.localeCompare(b.cab, 'ru');
    return byCab !== 0 ? byCab : a.art.localeCompare(b.art, 'ru');
  });

  // Остаток из U:Y
  const fboSumByKey = new Map();
  const lastDst = shDst.getLastRow();
  const stocksBody = lastDst > 1 ? shDst.getRange(2, 21, lastDst - 1, 5).getValues() : [];
  for (const r of stocksBody) {
    const cab = String(r[0] || '').replace(/\u00A0/g, ' ').trim();
    const art = String(r[1] || '').trim();
    const qty = _num0(r[2]);
    if (!cab || !art || qty <= 0) continue;
    fboSumByKey.set(cab + '|' + art, (fboSumByKey.get(cab + '|' + art) || 0) + qty);
  }

  // Скорости из K:S
  const speedsByKey = new Map();
  const toEdge = _edgesToEndOfYesterday().to;
  const wins = [7,14,21,28];

  if (lastDst > 1) {
    const ordBody = shDst.getRange(2, 11, lastDst - 1, 9).getValues(); // K:S
    for (const r of ordBody) {
      const cab = String(r[0] || '').replace(/\u00A0/g, ' ').trim();      // K
      const dt  = r[1] instanceof Date ? r[1] : (r[1] ? new Date(r[1]) : null);  // L
      const art = String(r[3] || '').trim();                               // N
      const qty = _num0(r[4]);                                             // O
      const method = String(r[5] || '').trim().toUpperCase();              // P
      if (!cab || !art || !dt || !isFinite(dt.getTime()) || qty <= 0) continue;

      const key = cab + '|' + art;
      if (!speedsByKey.has(key)) {
        speedsByKey.set(key, {
          w7:{sum:0,hasFbs:false}, w14:{sum:0,hasFbs:false},
          w21:{sum:0,hasFbs:false}, w28:{sum:0,hasFbs:false}
        });
      }
      const rec = speedsByKey.get(key);

      for (const w of wins) {
        const wEdge = _windowEdges(w, toEdge);
        if (dt >= wEdge.since && dt <= wEdge.to) {
          const slot = w===7?'w7':w===14?'w14':w===21?'w21':'w28';
          rec[slot].sum += qty;
          if (method === 'FBS') rec[slot].hasFbs = true; // совместимость
        }
      }
    }
  }

  // Формирование A:I
  const out = [];
  const fmts = [];
  const FMT_PLAIN = '0.00';
  const FMT_BRACK = '"["0.00"]"';

  for (const {cab, art} of keyList) {
    const key = cab + '|' + art;

    const fboLeftNum = _num0(fboSumByKey.get(key) || 0);
    const fboLeft = fboLeftNum === 0 ? '' : fboLeftNum;

    const rec = speedsByKey.get(key) || {
      w7:{sum:0,hasFbs:false}, w14:{sum:0,hasFbs:false},
      w21:{sum:0,hasFbs:false}, w28:{sum:0,hasFbs:false}
    };
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
    _applyCabinetRowFills_WB(shDst, out, 1, 9);

    shDst.getRange(2, 1, out.length, 2).setHorizontalAlignment('left').setVerticalAlignment('middle');
    shDst.getRange(2, 3, out.length, 7).setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  SpreadsheetApp.flush();
  ss.toast('Расчет (WB): записано строк ' + out.length, 'Готово', 4);
}

/* =========================
 *          fiz2_WB
 * ========================= */
function fiz2_WB() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_FIZ_WB) || ss.insertSheet(SHEET_FIZ_WB);
  _ensureFizHeader_WB(sh);

  const edges = _getLast30DaysEdgesExclToday(); // {sinceISO,toISO}
  const pick  = _getSelectedCabinetsFromParamsFIZ_WB(); // {selected,total,cabinetsToProcess,tokensByCab}

  const out = []; // K:S
  const okSet = new Set();

  for (const cab of pick.cabinetsToProcess) {
    const tokens = pick.tokensByCab.get(cab) || [];
    if (!tokens.length) continue;

    let orders = [];
    for (let i=0;i<tokens.length;i++) {
      try {
        const api = new WB(tokens[i]);
        orders = api.getOrders({ dateFrom: edges.sinceISO }) || [];
        break;
      } catch (e) {
        if (i === tokens.length - 1) console.log(`WB.getOrders fail for ${cab}: ${e && e.message || e}`);
      }
    }

    const before = out.length;
    for (const row of (orders || [])) {
      if (row.isCancel) continue;

      const dt = row.date ? new Date(row.date) : null;
      if (!(dt && isFinite(dt.getTime()))) continue;
      if (!(dt >= new Date(edges.sinceISO) && dt <= new Date(edges.toISO))) continue;

      const offer = String(row.supplierArticle || '').trim();
      if (!offer) continue;

      const qty = 1;
      const wh  = String(row.warehouseName || '').trim();
      const postingId = String(row.srid || row.orderId || row.id || '').trim();

      // K:S = Каб, Дата, ID отправления, Артикул, Кол-во, Метод, Склад отгр., Кл. отгр., Кл. дост.
      out.push([ cab, dt, postingId, offer, qty, 'WB', wh || '', '', '' ]);
    }
    if (out.length > before) okSet.add(cab);
  }

  _clearBlock(sh, 2, 11, 99999, 9); // K:S
  if (out.length) {
    const rng = sh.getRange(2, 11, out.length, 9);
    rng.setValues(out);

    sh.getRange(2, 11, out.length, 9)
      .setFontFamily('Roboto').setFontSize(8).setFontColor('#000000')
      .setFontWeight('normal').setHorizontalAlignment('left').setVerticalAlignment('middle');

    rng.sort([{ column: 12, ascending: false }]); // по дате ↓ (L)
  }

  SpreadsheetApp.flush();
  ss.toast('Заказы (WB): записано строк ' + out.length, 'Готово', 3);

  return { rows: out.length, okCabs: Array.from(okSet) };
}

/* =========================
 *          fiz3_WB
 * ========================= */
function fiz3_WB() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_FIZ_WB) || ss.insertSheet(SHEET_FIZ_WB);
  _ensureFizHeader_WB(sh);

  const pick = _getSelectedCabinetsFromParamsFIZ_WB(); // с токенами

  const out = []; // U:Y
  const okSet = new Set();

  for (const cab of pick.cabinetsToProcess) {
    const tokens = pick.tokensByCab.get(cab) || [];
    if (!tokens.length) continue;

    let stocks = [];
    for (let i=0;i<tokens.length;i++) {
      try {
        const api = new WB(tokens[i]);
        stocks = api.getStocks() || [];
        break;
      } catch (e) {
        if (i === tokens.length - 1) console.log(`WB.getStocks fail for ${cab}: ${e && e.message || e}`);
      }
    }

    const before = out.length;
    for (const it of (stocks || [])) {
      const offer = String(it.supplierArticle || '').trim();
      const qty = _num0(it.quantity);
      if (!offer || qty <= 0) continue;

      const whName  = String(it.warehouseName || it.warehouse || '').trim();
      const cluster = ''; // WB не отдаёт код кластера в этом API

      out.push([cab, offer, qty, whName, cluster]);
    }
    if (out.length > before) okSet.add(cab);
  }

  _clearBlock(sh, 2, 21, 99999, 5); // U:Y
  if (out.length) {
    sh.getRange(2, 21, out.length, 5).setValues(out);
    sh.getRange(2, 21, out.length, 5)
      .setFontFamily('Roboto').setFontSize(8).setFontColor('#000000')
      .setFontWeight('normal').setHorizontalAlignment('left').setVerticalAlignment('middle');
  }

  SpreadsheetApp.flush();
  ss.toast('Остатки (WB): записано строк ' + out.length, 'Готово', 3);

  return { rows: out.length, okCabs: Array.from(okSet) };
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

/* Хедеры и стили */
function _ensureFizHeader_WB(sh) {
  // --- A:I ---
  const headersAI = ['[ WB ] Кабинет','Артикул','Остаток','В поставке','Скорость','Ск [ 7 ]','Ск [ 14 ]','Ск [ 21 ]','Ск [ 28 ]'];
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

  // жёлтый тег "[ WB ]"
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

/* Построчные заливки по Кабинету (с нормализацией NBSP) */
function _applyCabinetRowFills_WB(sh, rows, startCol, numCols) {
  try {
    const colorMap = _readCabinetColorMapWB_();
    if (!rows || !rows.length) return;
    const fills = [];
    for (let i = 0; i < rows.length; i++) {
      const cab = String(rows[i][0] || '').replace(/\u00A0/g, ' ').trim();
      const color = (colorMap && colorMap.get(cab)) || '#ffffff';
      const line = new Array(numCols);
      for (let c = 0; c < numCols; c++) line[c] = color;
      fills.push(line);
    }
    sh.getRange(2, startCol, rows.length, numCols).setBackgrounds(fills);
  } catch (_) {}
}

/* Универсальные стили данных */
function _applyFizDataStyles(sh, startRow, numRows, numCols, fontSize /* =8 */) {
  if (numRows <= 0) return;
  const size = (fontSize && +fontSize) ? +fontSize : 8;
  sh.getRange(startRow, 1, numRows, numCols)
    .setFontFamily('Roboto').setFontSize(size).setFontWeight('normal').setFontColor('#000000')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
}

/* Чтение выбранных WB-кабинетов из «⚙️ Параметры» (с фолбэком) */
function _getSelectedCabinetsFromParamsFIZ_WB() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_PARAMS_WB);

  const res = { selected: [], total: 0, cabinetsToProcess: [], tokensByCab: new Map() };

  if (!sh) {
    const arts = ss.getSheetByName(SHEET_ARTS_WB);
    if (arts && arts.getLastRow() > 1) {
      const names = arts.getRange(2,1,arts.getLastRow()-1,1).getDisplayValues()
        .map(r => String(r[0]||'').trim())
        .filter(Boolean);
      const uniq = Array.from(new Set(names));
      res.selected = uniq.slice();
      res.total = uniq.length;
      res.cabinetsToProcess = uniq.slice();
    }
    return res;
  }

  const last = sh.getLastRow();
  if (last <= 1) return res;

  const rows  = sh.getRange(2, 1, last - 1, 8).getDisplayValues(); // A..H
  for (let i = 0; i < rows.length; i++) {
    const name  = String(rows[i][0] || '').replace(/\u00A0/g, ' ').trim(); // A
    const token = String(rows[i][2] || '').trim();                         // C
    const platf = String(rows[i][3] || '').trim().toUpperCase();           // D
    const on    = String(rows[i][7] || '').trim().toUpperCase();           // H
    const isOn  = (on === 'TRUE' || on === 'ИСТИНА' || on === 'ДА');

    if (!name || !token || !(platf === 'WILDBERRIES' || platf === 'WB')) continue;

    if (!res.tokensByCab.has(name)) res.tokensByCab.set(name, []);
    const arr = res.tokensByCab.get(name);
    if (arr.indexOf(token) === -1) arr.push(token);
    if (isOn && res.selected.indexOf(name) === -1) res.selected.push(name);
  }

  const uniqAll = Array.from(res.tokensByCab.keys());
  const uniqSel = Array.from(new Set(res.selected));
  res.total     = uniqAll.length;
  res.selected  = uniqSel;
  res.cabinetsToProcess = (!uniqSel.length || uniqSel.length === uniqAll.length) ? uniqAll : uniqSel;

  return res;
}

/* ---------- Безопасные фолбэки, если REF не прогрузился ---------- */
const SHEET_FIZ_WB     = (typeof REF !== 'undefined' && REF.SHEETS && REF.SHEETS.FIZ_WB)     ? REF.SHEETS.FIZ_WB     : '[WB] Физ. оборот';
const SHEET_ARTS_WB    = (typeof REF !== 'undefined' && REF.SHEETS && REF.SHEETS.ARTS_WB)    ? REF.SHEETS.ARTS_WB    : '[WB] Артикулы';
const SHEET_PARAMS_WB  = (typeof REF !== 'undefined' && REF.SHEETS && REF.SHEETS.PARAMS)     ? REF.SHEETS.PARAMS     : '⚙️ Параметры';

function _readCabinetColorMapWB_() {
  try {
    if (typeof REF !== 'undefined' && typeof REF.readCabinetColorMap === 'function') {
      return REF.readCabinetColorMap('WILDBERRIES');
    }
  } catch (_) {}
  return new Map();
}
