/** =========================================================
 * getFIZICAL.gs — единый файл для OZON/WB
 *  Публичные запускатели:
 *    - getFiz_OZ()
 *    - getFiz_WB()
 *  Макеты листов:
 *    A:D — [TAG] Кабинет | Артикул | Остаток | Скорость
 *    F:N — Заказы 30д (до конца вчера): 9 колонок
 *    P:T — Остатки (>0): 5 колонок
 * ========================================================= */

/* ---------- Безопасные резолверы ---------- */
function _fiz_sheetName(key, fallback) {
  try { return REF.sheetName(key, fallback); } catch (_) { return fallback; }
}
function _fiz_colorMap(platformTag) {
  try { return REF.readCabinetColorMap(platformTag); } catch (_) { return new Map(); }
}

/* ---------- Публичные ---------- */
function getFiz_OZ() {
  const ss = SpreadsheetApp.getActive();
  const sheetTitle = _fiz_sheetName('FIZ_OZ','[OZ] Физ. оборот');
  const sh = ss.getSheetByName(sheetTitle) || ss.insertSheet(sheetTitle);

  _fiz_prepareHeaderOnce(sh, 'OZ'); // шапка, ширины, заморозка

  // 1) Остатки → P:T
  const rStocks = _fiz_fillStocks_OZ(sh);
  ss.toast('Остатки: ' + (rStocks.rows||0) + ' строк; кабинетов: ' + (rStocks.okCabs||[]).length, sheetTitle, 3);

  // 2) Заказы → F:N (30 дней до конца вчера)
  const rOrders = _fiz_fillOrders_OZ(sh);
  ss.toast('Заказы (30д): ' + (rOrders.rows||0) + ' строк; кабинетов: ' + (rOrders.okCabs||[]).length, sheetTitle, 3);

  // 3) Основной блок A:D по «Артикулам» + вычисление скоростей и остатков
  const rows = _fiz_buildMainBlockFromRefs(sh, 'OZ');
  ss.toast('Скорость рассчитана для ' + rows + ' артикулов', sheetTitle, 3);

  // Лог кабинетов в ⚙️Параметры
  try {
    const seen = new Set(); const ok = [];
    for (const arr of [rStocks.okCabs||[], rOrders.okCabs||[]]) {
      for (const c of arr) if (!seen.has(c)) { seen.add(c); ok.push(c); }
    }
    REF.logRun('Физ.оборот OZ', ok, 'OZON');
  } catch (_){}

  SpreadsheetApp.flush();
}

function getFiz_WB() {
  const ss = SpreadsheetApp.getActive();
  const sheetTitle = _fiz_sheetName('FIZ_WB','[WB] Физ. оборот');
  const sh = ss.getSheetByName(sheetTitle) || ss.insertSheet(sheetTitle);

  _fiz_prepareHeaderOnce(sh, 'WB'); // шапка, ширины, заморозка

  // 1) Остатки → P:T
  const rStocks = _fiz_fillStocks_WB(sh);
  ss.toast('Остатки: ' + (rStocks.rows||0) + ' строк; кабинетов: ' + (rStocks.okCabs||[]).length, sheetTitle, 3);

  // 2) Заказы → F:N
  const rOrders = _fiz_fillOrders_WB(sh);
  ss.toast('Заказы (30д): ' + (rOrders.rows||0) + ' строк; кабинетов: ' + (rOrders.okCabs||[]).length, sheetTitle, 3);

  // 3) Основной блок A:D
  const rows = _fiz_buildMainBlockFromRefs(sh, 'WB');
  ss.toast('Скорость рассчитана для ' + rows + ' артикулов', sheetTitle, 3);

  // Лог кабинетов
  try {
    const seen = new Set(); const ok = [];
    for (const arr of [rStocks.okCabs||[], rOrders.okCabs||[]]) {
      for (const c of arr) if (!seen.has(c)) { seen.add(c); ok.push(c); }
    }
    REF.logRun('Физ.оборот WB', ok, 'WILDBERRIES');
  } catch (_){}

  SpreadsheetApp.flush();
}

/* =========================================================
 *                    Х Е Л П Е Р Ы
 * ========================================================= */

/* ---------- Шапка/ширины/цвета (A:D + F:N + P:T) ---------- */
function _fiz_prepareHeaderOnce(sh, platform /* 'OZ'|'WB' */) {
  // очищаем строку 1 целиком A:T и задаём единообразные стили
  sh.getRange(1, 1, 1, 20).clear({ contentsOnly: true });
  sh.getRange(1, 1, 1, 20)
    .setBackground(null).setFontColor('#000')
    .setFontFamily('Roboto').setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setWrap(false);
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 22);

  // --- A:D ---
  const headerTag = (platform === 'OZ') ? '[ OZ ] Кабинет' : '[ WB ] Кабинет';
  const H = [headerTag, 'Артикул', 'Остаток', 'Скорость'];

  sh.getRange(1, 1, 1, 4).setValues([H]).setFontColor('#ffffff');
  sh.getRange(1, 1, 1, 2).setBackground('#434343'); // A:B
  sh.getRange(1, 3, 1, 1).setBackground('#38761d'); // C «Остаток»
  sh.getRange(1, 4, 1, 1).setBackground('#1155cc'); // D «Скорость»

  // Тег в A1 — жёлтым
  try {
    const a1Text = H[0];
    const a1 = sh.getRange(1,1);
    const tagEnd = a1Text.indexOf(']') + 1;
    const b = SpreadsheetApp.newRichTextValue().setText(a1Text);
    b.setTextStyle(0, tagEnd, SpreadsheetApp.newTextStyle().setForegroundColor('#ffff00').setBold(true).build());
    b.setTextStyle(tagEnd, a1Text.length, SpreadsheetApp.newTextStyle().setForegroundColor('#ffffff').setBold(true).build());
    a1.setRichTextValue(b.build());
  } catch (_) {}

  // --- F:N (Заказы) ---
  const ORD = ['Кабинет','Дата','ID отправления','Артикул','Количество','Метод','Склад отгрузки','Кластер отгрузки','Кластер доставки'];
  sh.getRange(1, 6, 1, ORD.length).setValues([ORD]).setFontColor('#ffffff').setBackground('#434343');

  // --- P:T (Остатки) ---
  const STK = ['Кабинет','Артикул','Количество','Склад хранения','Кластер хранения'];
  sh.getRange(1, 16, 1, STK.length).setValues([STK]).setFontColor('#ffffff').setBackground('#434343');

  // --- ширины: A,B авто+50; C,D = 90; остальное не трогаем
  _fiz_autoPlus(sh, 1, 50);
  _fiz_autoPlus(sh, 2, 50);
  sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 90);
}

/* автоширина + добавка (с flush, чтобы не зависало старое значение) */
function _fiz_autoPlus(sh, col, plus) {
  sh.autoResizeColumn(col);
  SpreadsheetApp.flush();
  const w = sh.getColumnWidth(col);
  sh.setColumnWidth(col, w + (plus || 0));
}

/* ---------- Общие утилиты ---------- */
function _fiz_num(v){ return isFinite(+v) ? (+v) : 0; }

function _fiz_edges30dExclToday() {
  const now = new Date();
  const to = new Date(now);
  to.setDate(to.getDate() - 1);
  to.setHours(23,59,59,999);
  const since = new Date(to);
  since.setDate(since.getDate() - 29);
  since.setHours(0,0,0,0);
  return { sinceISO: since.toISOString(), toISO: to.toISOString(), to };
}

function _fiz_windowEdges(days, to) {
  const since = new Date(to);
  since.setDate(since.getDate() - (days - 1));
  since.setHours(0,0,0,0);
  return { since, to };
}

function _fiz_clear(sh, r, c, h, w) {
  sh.getRange(r, c, h, w).clear({ contentsOnly: true });
}

/* Заполнить заливки строк по цветам кабинетов (A:D) */
function _fiz_applyCabinetFills(sh, rows, startCol, numCols, platformTag /* 'OZON'|'WILDBERRIES' */) {
  try {
    const cmap = _fiz_colorMap(platformTag);
    if (!rows || !rows.length) return;
    const fills = [];
    for (let i=0;i<rows.length;i++){
      const cab = String(rows[i][0]||'').replace(/\u00A0/g,' ').trim();
      const color = (cmap && cmap.get(cab)) || '#ffffff';
      const line = new Array(numCols).fill(color);
      fills.push(line);
    }
    sh.getRange(2, startCol, rows.length, numCols).setBackgrounds(fills);
  } catch(_){}
}

/* Читать выбранные кабинеты из ⚙️Параметры (универсально) */
function _fiz_readSelectedCabinets(platform /* 'OZ'|'WB' */) {
  const ss = SpreadsheetApp.getActive();
  const paramsName = _fiz_sheetName('PARAMS', '⚙️ Параметры');
  const sh = ss.getSheetByName(paramsName);

  const out = { selected: [], total: 0, all: [], tokensByCab: new Map(), /* WB only */ };
  const artsSheet = ss.getSheetByName(platform==='OZ' ? _fiz_sheetName('ARTS_OZ','[OZ] Артикулы') : _fiz_sheetName('ARTS_WB','[WB] Артикулы'));

  if (!sh) {
    if (artsSheet && artsSheet.getLastRow() > 1) {
      const names = artsSheet.getRange(2,1,artsSheet.getLastRow()-1,1).getDisplayValues()
        .map(r => String(r[0]||'').trim()).filter(Boolean);
      const uniq = Array.from(new Set(names));
      out.selected = uniq.slice();
      out.total = uniq.length;
      out.all = uniq.slice();
    }
    return out;
  }

  const last = sh.getLastRow();
  if (last <= 1) return out;

  const rows  = sh.getRange(2, 1, last - 1, 8).getDisplayValues(); // A..H
  for (let i=0;i<rows.length;i++) {
    const name  = String(rows[i][0] || '').replace(/\u00A0/g, ' ').trim(); // A
    const token = String(rows[i][2] || '').trim();                         // C
    const platf = String(rows[i][3] || '').trim().toUpperCase();           // D
    const on    = String(rows[i][7] || '').trim().toUpperCase();           // H
    const isOn  = (on === 'TRUE' || on === 'ИСТИНА' || on === 'ДА');

    if (!name) continue;
    if (platform === 'OZ' && !(platf === 'OZON' || platf === 'OZ')) continue;
    if (platform === 'WB' && !(platf === 'WILDBERRIES' || platf === 'WB')) continue;

    if (platform === 'WB' && token) {
      if (!out.tokensByCab.has(name)) out.tokensByCab.set(name, []);
      const arr = out.tokensByCab.get(name);
      if (arr.indexOf(token) === -1) arr.push(token);
    }

    out.all.push(name);
    if (isOn && out.selected.indexOf(name) === -1) out.selected.push(name);
  }

  const uniqAll = Array.from(new Set(out.all));
  const uniqSel = Array.from(new Set(out.selected));
  out.total     = uniqAll.length;
  out.selected  = uniqSel;
  out.all       = uniqAll;
  return out;
}

/* ---------- A:D построение по справочнику артикулов + F:N + P:T ---------- */
function _fiz_buildMainBlockFromRefs(sh, platform /* 'OZ'|'WB' */) {
  const ss = SpreadsheetApp.getActive();
  const artsName = (platform==='OZ') ? _fiz_sheetName('ARTS_OZ','[OZ] Артикулы') : _fiz_sheetName('ARTS_WB','[WB] Артикулы');
  const shArts = ss.getSheetByName(artsName);
  if (!shArts) throw new Error('Нет листа «' + artsName + '»');

  // пары (Кабинет|Артикул) из артикулов, с учётом выбранных кабинетов
  const pick = _fiz_readSelectedCabinets(platform);
  const filterActive = pick.selected.length && pick.selected.length < pick.total;

  const artsLast = shArts.getLastRow();
  const artsRows = artsLast > 1 ? shArts.getRange(2, 1, artsLast - 1, 2).getDisplayValues() : [];
  const keyList = [], keySet = new Set();
  for (const r of artsRows) {
    const cab = String(r[0]||'').replace(/\u00A0/g,' ').trim();
    const art = String(r[1]||'').trim();
    if (!cab || !art) continue;
    if (filterActive && pick.selected.indexOf(cab) === -1) continue;
    const key = cab + '|' + art;
    if (!keySet.has(key)) { keySet.add(key); keyList.push({ cab, art }); }
  }
  keyList.sort((a,b) => (a.cab.localeCompare(b.cab,'ru') || a.art.localeCompare(b.art,'ru')));

  const last = sh.getLastRow();

  // Остаток из P:T
  const fboLeftByKey = new Map();
  const stocks = last>1 ? sh.getRange(2,16,last-1,5).getValues() : []; // P:T
  for (const r of stocks) {
    const cab = String(r[0]||'').replace(/\u00A0/g,' ').trim();
    const art = String(r[1]||'').trim();
    const qty = _fiz_num(r[2]);
    if (!cab || !art || qty <= 0) continue;
    const k = cab+'|'+art;
    fboLeftByKey.set(k, (fboLeftByKey.get(k)||0) + qty);
  }

  // Скорости из F:N — окна и MAX по FBO/FBS
  const wins=[7,14,21,28];
  const toEdge = _fiz_edges30dExclToday().to;
  const spFBO = new Map(), spFBS = new Map();
  if (last>1) {
    const ord = sh.getRange(2,6,last-1,9).getValues(); // F:N
    for (const r of ord) {
      const cab = String(r[0]||'').replace(/\u00A0/g,' ').trim();
      const dt  = r[1] instanceof Date ? r[1] : (r[1]?new Date(r[1]):null);
      const art = String(r[3]||'').trim();
      const qty = _fiz_num(r[4]);
      const method = String(r[5]||'').trim().toUpperCase();
      if (!cab || !art || !dt || !isFinite(dt.getTime()) || qty<=0) continue;

      const key=cab+'|'+art;
      const target = (method==='FBO') ? spFBO : (method==='FBS') ? spFBS : null;
      if (!target) continue;

      if (!target.has(key)) target.set(key,{w7:0,w14:0,w21:0,w28:0});
      for (const d of wins) {
        const e = _fiz_windowEdges(d, toEdge);
        if (dt>=e.since && dt<=e.to) {
          const slot=(d===7?'w7':d===14?'w14':d===21?'w21':'w28');
          target.get(key)[slot] += qty;
        }
      }
    }
  }
  const maxNorm = (rec) => Math.max((rec.w7||0)/7,(rec.w14||0)/14,(rec.w21||0)/21,(rec.w28||0)/28,0);

  // Формируем A:D
  const out=[], fmts=[];
  for (const {cab,art} of keyList) {
    const key=cab+'|'+art;
    const left=_fiz_num(fboLeftByKey.get(key)||0);
    const vFBO = spFBO.has(key)?maxNorm(spFBO.get(key)):0;
    const vFBS = spFBS.has(key)?maxNorm(spFBS.get(key)):0;
    const speed = vFBO + vFBS;

    out.push([cab, art, (left===0?'':left), (speed>0?Number(speed.toFixed(2)):'')]);
    fmts.push(['@','@','0.########','0.00']);
  }

  _fiz_clear(sh, 2, 1, 999999, 4); // очистить A:D
  if (out.length) {
    sh.getRange(2,1,out.length,4).setValues(out);
    sh.getRange(2,1,out.length,4).setNumberFormats(fmts);
    // выравнивание
    sh.getRange(2,3,out.length,2).setHorizontalAlignment('center').setVerticalAlignment('middle');

    // стили строк + заливка по кабинетам
    sh.getRange(2,1,out.length,4).setFontFamily('Roboto').setFontSize(10).setFontColor('#000').setFontWeight('normal');
    _fiz_applyCabinetFills(sh, out, 1, 4, (platform==='OZ'?'OZON':'WILDBERRIES'));
  }
  return out.length;
}

/* ---------- Заказы / Остатки: OZ ---------- */
function _fiz_fillOrders_OZ(sh) {
  const edges = _fiz_edges30dExclToday();
  const pick = _fiz_readSelectedCabinets('OZ');

  // Верифицируем список по аккаунтам OZONAPI (если есть)
  let accs = {};
  try { accs = OZONAPI.getAccounts() || {}; } catch(_) { accs = {}; }
  const valid = new Set(Object.keys(accs||{}));
  const toProcess = pick.selected.length ? pick.selected : pick.all;
  const cabs = toProcess.filter(c => valid.size ? valid.has(c) : true);

  const out=[], ok=new Set();

  for (const cab of cabs) {
    if (valid.size && !valid.has(cab)) continue;
    let fbo=[], fbs=[];
    try { fbo = new OZONAPI(cab).getOrdersFbo(30) || []; } catch(e){ fbo=[]; }
    try { fbs = new OZONAPI(cab).getOrdersFbs(30) || []; } catch(e){ fbs=[]; }

    const pushPosting = (arr, method) => {
      for (const p of (arr||[])) {
        const dtStr = p.in_process_at || p.shipment_date || p.shipped_at || p.created_at;
        if (!dtStr) continue;
        const dt = new Date(dtStr);
        if (!(dt >= new Date(edges.sinceISO) && dt <= new Date(edges.toISO))) continue;

        const statusRaw = String(p.status || p.state || '').trim().toLowerCase();
        if (statusRaw === 'cancelled' || statusRaw === 'canceled' || /^cancel/.test(statusRaw)) continue;

        const postingId = String(p.posting_number || p.postingNumber || p.id || '').trim();
        const analytics = p.analytics_data || {};
        const whName = analytics.warehouse_name || analytics.warehouse || '';

        let codeFrom = '', codeTo = '';
        const fd = p.financial_data || {};
        if (fd) {
          codeFrom = String(fd.cluster_from || '').trim();
          codeTo   = String(fd.cluster_to || '').trim();
        }

        const products = Array.isArray(p.products) ? p.products : (Array.isArray(p.resulting_products)?p.resulting_products:[]);
        for (const pr of (products||[])) {
          const offer = pr.offer_id || pr.offer || '';
          const qty   = _fiz_num(pr.quantity || pr.qty);
          if (!offer || qty<=0) continue;
          out.push([cab, dt, postingId, offer, qty, method, whName||'', codeFrom, codeTo]);
        }
      }
    };

    pushPosting(fbo, 'FBO');
    pushPosting(fbs, 'FBS');
    if (out.length) ok.add(cab);
  }

  _fiz_clear(sh, 2, 6, 999999, 9); // F:N
  if (out.length) {
    const rng = sh.getRange(2,6,out.length,9);
    rng.setValues(out);
    rng.setFontFamily('Roboto').setFontSize(8).setFontColor('#000').setFontWeight('normal')
       .setHorizontalAlignment('left').setVerticalAlignment('middle');
    rng.sort([{ column: 7, ascending: false }]); // сортируем по G (дата) ↓
  }
  return { rows: out.length, okCabs: Array.from(ok) };
}

function _fiz_fillStocks_OZ(sh) {
  const ss = SpreadsheetApp.getActive();
  const shArts = ss.getSheetByName(_fiz_sheetName('ARTS_OZ','[OZ] Артикулы'));
  if (!shArts) return { rows:0, okCabs:[] };

  const pick = _fiz_readSelectedCabinets('OZ');
  const filterActive = pick.selected.length && pick.selected.length < pick.total;

  // offers by cabinet из «[OZ] Артикулы»
  const offersByCab = new Map();
  const lastA = shArts.getLastRow();
  const rowsA = lastA>1 ? shArts.getRange(2,1,lastA-1,2).getDisplayValues() : [];
  for (const r of rowsA) {
    const cab = String(r[0]||'').trim();
    const offer = String(r[1]||'').trim();
    if (!cab || !offer) continue;
    if (filterActive && pick.selected.indexOf(cab) === -1) continue;
    if (!offersByCab.has(cab)) offersByCab.set(cab, new Set());
    offersByCab.get(cab).add(offer);
  }

  let accs = {};
  try { accs = OZONAPI.getAccounts() || {}; } catch(_) { accs = {}; }
  const out=[], ok=new Set();

  for (const [cab, offerSet] of offersByCab) {
    if (accs && !accs[cab]) continue;

    const api = new OZONAPI(cab);
    const offers = Array.from(offerSet);
    if (!offers.length) continue;

    let offerToSku = {};
    try { offerToSku = api.getSkusByOffers(offers) || {}; } catch(e){ offerToSku = {}; }

    const skus = Object.values(offerToSku).filter(Boolean);
    if (!skus.length) continue;

    let items = [];
    try { items = api.analyticsStocksBySkus(skus) || []; } catch(e){ items=[]; }

    const skuToOffer = {};
    for (const off in offerToSku) {
      const sku = offerToSku[off];
      if (sku != null) skuToOffer[String(sku)] = off;
    }

    const before = out.length;
    for (const it of (items||[])) {
      const sku = it && (it.sku ?? it.sku_id ?? it.id);
      const offer = (sku != null ? skuToOffer[String(sku)] : (it.offer_id || it.offer)) || '';
      if (!offer) continue;
      const qty = _fiz_num(it.available_stock_count || it.available || it.qty);
      if (qty <= 0) continue;
      const whName  = String(it.warehouse_name || it.warehouse || '').trim();
      const cluster = String(it.cluster_name || it.cluster || '').trim();
      out.push([cab, offer, qty, whName, cluster]);
    }
    if (out.length > before) ok.add(cab);
  }

  _fiz_clear(sh, 2, 16, 999999, 5); // P:T
  if (out.length) {
    sh.getRange(2,16,out.length,5).setValues(out);
    sh.getRange(2,16,out.length,5)
      .setFontFamily('Roboto').setFontSize(8).setFontColor('#000')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  }
  return { rows: out.length, okCabs: Array.from(ok) };
}

/* ---------- Заказы / Остатки: WB ---------- */
function _fiz_fillOrders_WB(sh) {
  const edges = _fiz_edges30dExclToday();
  const pick = _fiz_readSelectedCabinets('WB');

  const out=[], ok=new Set();

  for (const cab of (pick.selected.length ? pick.selected : pick.all)) {
    const tokens = pick.tokensByCab.get(cab) || [];
    if (!tokens.length) continue;

    let orders = [];
    for (let i=0;i<tokens.length;i++) {
      try { orders = new WB(tokens[i]).getOrders({ dateFrom: edges.sinceISO }) || []; break; }
      catch(e){ if (i===tokens.length-1) console.log('WB.getOrders fail for', cab, e && e.message || e); }
    }

    const before = out.length;
    for (const row of (orders||[])) {
      if (row.isCancel) continue;
      const dt = row.date ? new Date(row.date) : null;
      if (!(dt && isFinite(dt.getTime()))) continue;
      if (!(dt >= new Date(edges.sinceISO) && dt <= new Date(edges.toISO))) continue;

      const offer = String(row.supplierArticle||'').trim(); if (!offer) continue;
      const qty = 1; // WB ордеры как штучные события
      const wh  = String(row.warehouseName||'').trim();
      const postingId = String(row.srid || row.orderId || row.id || '').trim();

      out.push([ cab, dt, postingId, offer, qty, 'FBO', wh || '', '', '' ]);
    }
    if (out.length > before) ok.add(cab);
  }

  _fiz_clear(sh, 2, 6, 999999, 9); // F:N
  if (out.length) {
    const rng = sh.getRange(2,6,out.length,9);
    rng.setValues(out);
    rng.setFontFamily('Roboto').setFontSize(8).setFontColor('#000').setFontWeight('normal')
       .setHorizontalAlignment('left').setVerticalAlignment('middle');
    rng.sort([{ column: 7, ascending: false }]); // по G ↓ (дата)
  }
  return { rows: out.length, okCabs: Array.from(ok) };
}

function _fiz_fillStocks_WB(sh) {
  const pick = _fiz_readSelectedCabinets('WB');
  const out=[], ok=new Set();

  for (const cab of (pick.selected.length ? pick.selected : pick.all)) {
    const tokens = pick.tokensByCab.get(cab) || [];
    if (!tokens.length) continue;

    let stocks=[];
    for (let i=0;i<tokens.length;i++) {
      try { stocks = new WB(tokens[i]).getStocks() || []; break; }
      catch(e){ if (i===tokens.length-1) console.log('WB.getStocks fail for', cab, e && e.message || e); }
    }

    const before = out.length;
    for (const it of (stocks||[])) {
      const offer = String(it.supplierArticle||'').trim();
      const qty = _fiz_num(it.quantity);
      if (!offer || qty<=0) continue;
      const wh = String(it.warehouseName||it.warehouse||'').trim();
      out.push([cab, offer, qty, wh, '']);
    }
    if (out.length>before) ok.add(cab);
  }

  _fiz_clear(sh, 2, 16, 999999, 5); // P:T
  if (out.length) {
    sh.getRange(2,16,out.length,5).setValues(out);
    sh.getRange(2,16,out.length,5)
      .setFontFamily('Roboto').setFontSize(8).setFontColor('#000')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  }
  return { rows: out.length, okCabs: Array.from(ok) };
}
