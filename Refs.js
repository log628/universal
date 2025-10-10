/** =========================================================
 * Refs.gs — единый справочник и утилиты для проекта (OZ/WB)
 *  - Имена листов и именованные диапазоны — централизованы в REF.SHEETS / REF.*
 *  - Безопасные резолверы имён листов: REF.sheetName(key, fallback)
 *  - Справочники заголовков и ширин для «[OZ]/[WB] Физ. оборот»
 *  - Числовые парсеры, логгер, чтение настроек «Мин. запас на карточку»
 *
 * Требование загрузки: положите Refs.gs выше в списке файлов, чтобы REF был
 * доступен для остальных *.gs (или используйте безопасные фолбэки из REF.*)
 * ========================================================= */

var REF = (function () {
  var REF = {};

  /* =========================
   *        Л И С Т Ы
   * ========================= */
  REF.SHEETS = {
    // Каталоги артикулов
    ARTS_OZ: '[OZ] Артикулы',
    ARTS_WB: '[WB] Артикулы',

    // Физический оборот (рабочие листы с расчётами)
    FIZ_OZ:  '[OZ] Физ. оборот',
    FIZ_WB:  '[WB] Физ. оборот',

    // Прочие общие листы
    PARAMS:  '⚙️ Параметры',
    RATES:   '🔖 Тарифы',
    SS:      '🍔 СС',
    CALC:    '⚖️ Калькулятор',

    // Форкаст
    FORECAST: '🎏 Форкаст'
  };

  /** Универсальный безопасный резолвер имени листа по ключу REF.SHEETS */
  REF.sheetName = function (key, fallback) {
    try { return (REF && REF.SHEETS && REF.SHEETS[key]) || fallback || String(key || ''); }
    catch (_) { return fallback || String(key || ''); }
  };

  // ✅ Контрол выбора кабинета — ИМЕНОВАННЫЙ ДИАПАЗОН
  REF.CTRL_RANGE_A1 = 'muff_cabs';

  /* =========================
   *  К У Л Д А У Н / З А Н Я Т О
   * ========================= */
  REF.COOLDOWN_MS = 5000; // 5 сек
  var DP = PropertiesService.getDocumentProperties();
  var KEY_BUSY  = 'ref_busy_flag_bool';
  var KEY_COOLT = 'ref_cooldown_last_ms';

  REF.isBusy        = function(){ return String(DP.getProperty(KEY_BUSY)||'') === '1'; };
  REF.busyStart     = function(){ DP.setProperty(KEY_BUSY, '1'); };
  REF.busyEnd       = function(){ DP.setProperty(KEY_BUSY, '');  };
  REF.cooldownStamp = function(){ DP.setProperty(KEY_COOLT, String(Date.now())); };
  REF.cooldownRemainMs = function(){
    var last = Number(DP.getProperty(KEY_COOLT)) || 0;
    return Math.max(0, REF.COOLDOWN_MS - (Date.now() - last));
  };

  /* =========================
   *      Х Е Д Е Р Ы  A:M
   * ========================= */
  REF.ARTS_HEADERS_BASE = [
    'Кабинет','Артикул','Отзывы','Рейтинг','Категория',
    'FBO','FBS','RFBS','Объем','Цена',
    'SKU','Раздел','Своя категория'
  ];
  REF.getArtsHeaders = function (tag /* 'OZ'|'WB'|string */) {
    var hdr = REF.ARTS_HEADERS_BASE.slice();
    var t = String(tag || '').trim().toUpperCase();
    if (t) hdr[0] = '[ ' + t + ' ] ' + 'Кабинет';
    return hdr;
  };
  REF.ARTS_COLS = { A:1,B:2,C:3,D:4,E:5,F:6,G:7,H:8,I:9,J:10,K:11,L:12,M:13 };
  REF.ARTS_TOTAL_COLS = 13;

  REF.ensureArtsLayout10 = function (sheetName, tag /* optional */) {
    if (!sheetName) throw new Error('ensureArtsLayout10: sheetName обязателен');
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    if (sh.getMaxRows() < 1) sh.insertRowsAfter(1, 1);

    var need = REF.ARTS_TOTAL_COLS;
    var cur  = sh.getMaxColumns();
    if (cur < need) sh.insertColumnsAfter(cur, need - cur);
    if (cur > need) sh.deleteColumns(need + 1, cur - need);

    var headers = tag ? REF.getArtsHeaders(tag) : REF.ARTS_HEADERS_BASE;
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  };
  REF.ensureArtsLayout12 = function (sheetName, tag) { REF.ensureArtsLayout10(sheetName, tag); };



REF.FIZ = {
  OZ: {
    SHEET_KEY: 'FIZ_OZ',
    TAG_LABEL: '[ OZ ] Кабинет',
    HEADERS_AI: [
      '[ OZ ] Кабинет',  // A
      'Артикул',         // B
      'Остаток',         // C  (FBO суммарно)
      'Скорость'         // D  (max FBO + max FBS)
    ],
    // ширины под 4 колонки
    COL_WIDTHS: { A: null, B: null, C: 90, D: 110 },
    NUM_FORMATS_ROW: ['@','@','0.########','0.00']
  },
  WB: {
    SHEET_KEY: 'FIZ_WB',
    TAG_LABEL: '[ WB ] Кабинет',
    HEADERS_AI: [
      '[ WB ] Кабинет',  // A
      'Артикул',         // B
      'Остаток',         // C
      'Скорость'         // D
    ],
    COL_WIDTHS: { A: null, B: null, C: 90, D: 110 },
    NUM_FORMATS_ROW: ['@','@','0.########','0.00']
  }
};


/* =========================
 *  «Физ. оборот»: координаты таблиц
 *  MAIN (A:D) — выгрузка: "[ TAG ] Кабинет" | "Артикул" | "Остаток" | "Скорость"
 *  ORDERS (F:N) — заказы, STOCKS (P:T) — остатки по складам
 * ========================= */
REF.FIZ.SECTIONS = {
  MAIN:   { startCol: 1,  width: 4,  headerA1: 'A1:D1', bodyA1: 'A2:D'  },
  ORDERS: { startCol: 6,  width: 9,  headerA1: 'F1:N1', bodyA1: 'F2:N'  },
  STOCKS: { startCol: 16, width: 5,  headerA1: 'P1:T1', bodyA1: 'P2:T'  }
};

/** Имя листа «Физ. оборот» по платформе */
REF.FIZ.sheetNameByPlatform = function (platform /* 'OZ'|'WB' */) {
  var p = String(platform||'').trim().toUpperCase();
  if (p === 'OZ') return REF.sheetName('FIZ_OZ', '[OZ] Физ. оборот');
  if (p === 'WB') return REF.sheetName('FIZ_WB', '[WB] Физ. оборот');
  return '';
};

/** A:D — заголовки для основной таблицы (берём платформенные) */
REF.FIZ.mainHeadersByPlatform = function (platform /* 'OZ'|'WB' */) {
  var p = String(platform||'').trim().toUpperCase();
  if (p === 'OZ') return (REF.FIZ.OZ && REF.FIZ.OZ.HEADERS_AI) ? REF.FIZ.OZ.HEADERS_AI : ['[ OZ ] Кабинет','Артикул','Остаток','Скорость'];
  if (p === 'WB') return (REF.FIZ.WB && REF.FIZ.WB.HEADERS_AI) ? REF.FIZ.WB.HEADERS_AI : ['[ WB ] Кабинет','Артикул','Остаток','Скорость'];
  return ['[ ? ] Кабинет','Артикул','Остаток','Скорость'];
};

/** Удобные геттеры диапазонов для MAIN (A:D) */
REF.FIZ.getMainHeaderRange = function (sheet) {
  var a1 = REF.FIZ.SECTIONS.MAIN.headerA1;
  return sheet.getRange(a1);
};
REF.FIZ.getMainBodyRange = function (sheet, lastRow /* optional */) {
  // Если lastRow не задан — вернём открытый A1-диапазон "A2:D" (для чтения/очистки)
  var a1 = REF.FIZ.SECTIONS.MAIN.bodyA1;
  if (!lastRow || lastRow < 2) return sheet.getRange(a1);
  // Иначе — точный прямоугольник по текущему lastRow
  var h = lastRow - 1;
  return sheet.getRange(2, REF.FIZ.SECTIONS.MAIN.startCol, h, REF.FIZ.SECTIONS.MAIN.width);
};

REF.FIZ.applyMainWidths = function (sheet, platform /* 'OZ'|'WB' optional */) {
  // A,B авто +50
  sheet.autoResizeColumn(1); sheet.setColumnWidth(1, sheet.getColumnWidth(1) + 50);
  sheet.autoResizeColumn(2); sheet.setColumnWidth(2, sheet.getColumnWidth(2) + 50);

  var p = String(platform||'').trim().toUpperCase();
  var cfg = (p==='WB' ? REF.FIZ.WB : REF.FIZ.OZ); // по умолчанию OZ
  var cw = (cfg && cfg.COL_WIDTHS) || {C:90, D:110};
  sheet.setColumnWidth(3, cw.C || 90);
  sheet.setColumnWidth(4, cw.D || 110);
};







  // Цвета заголовков для «Физ. оборот» (новый вид)
  REF.FIZ_COLORS = {
    HEADER_MAIN_BG: '#434343', // для «тёмной» шапки где нужно
    A_B_BG:         '#434343', // A:B
    C_BG:           '#38761d', // Остаток FBO
    D_E_BG:         '#1155cc', // Скорости
    F_G_BG:         '#6d9eeb', // Окна (расчётные «Ск [N]»)
    NEED_LEFT_BG:   '#741b47'  // Для заголовка «Потреб …» слева (если понадобится)
  };

  /* =========================
   *     Ч И С Л О В Ы Е
   * ========================= */
  REF.toNumber = function (v) {
    if (v == null) return 0;
    if (typeof v === 'number' && isFinite(v)) return v;
    var s = String(v).trim();
    if (!s) return 0;
    s = s.replace(/\u00A0|\u2007|\u202F|\u2009/g, '').replace(/\s+/g, '');
    s = s.replace(/[^0-9.,\-]/g, '');
    if (s.indexOf(',') > -1 && s.indexOf('.') > -1) s = s.replace(/\./g, '').replace(',', '.');
    else if (s.indexOf(',') > -1) s = s.replace(',', '.');
    var num = parseFloat(s);
    return isFinite(num) ? num : 0;
  };
  REF.round2 = function (n) { var x=Number(n); return isFinite(x) ? Math.round(x*100)/100 : 0; };
  REF.toComma = function (v) { return (v==null?'':String(v)).replace(/\./g, ','); };
  REF.toDot   = function (v) { return (v==null?'':String(v)).replace(/,/g, '.');  };

  /* =========================
   *     Т А Р И Ф Ы / Л О Г И С Т И К А
   * ========================= */
  REF.readTariffPercent = function (label) {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.RATES);
    if (!sh) return 0;
    var last = sh.getLastRow();
    if (last < 2) return 0;
    var vals = sh.getRange(2, 1, last - 1, 2).getDisplayValues(); // A:B
    for (var i=0;i<vals.length;i++){
      if (String(vals[i][0]||'').trim() === label) {
        var v = REF.toNumber(vals[i][1]);
        return isFinite(v) ? v : 0;
      }
    }
    return 0;
  };

  // D:G — Схема | Статья | До объема | Ставка
  REF.readRate = function (scheme, article) {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.RATES);
    if (!sh) return 0;
    var last = sh.getLastRow();
    if (last < 2) return 0;
    var vals = sh.getRange(2, 4, last - 1, 4).getDisplayValues(); // D:E:F:G
    for (var i=0;i<vals.length;i++){
      var sch = String(vals[i][0]||'').trim();
      var art = String(vals[i][1]||'').trim();
      if (sch === scheme && art === article) {
        var rate = REF.toNumber(vals[i][3]);
        return isFinite(rate) ? rate : 0;
      }
    }
    return 0;
  };

  // Пороги логистики { r1,r2,r3,r3p } по схеме/статье
  REF.readLogisticTiers = function (scheme, article) {
    var out = { r1:0, r2:0, r3:0, r3p:0 };
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.RATES);
    if (!sh) return out;
    var last = sh.getLastRow();
    if (last < 2) return out;

    var vals = sh.getRange(2, 4, last - 1, 4).getDisplayValues(); // D:E:F:G
    for (var i=0;i<vals.length;i++){
      var sch = String(vals[i][0]||'').trim();
      var art = String(vals[i][1]||'').trim();
      if (sch !== scheme || art !== article) continue;

      var tier = String(vals[i][2]||'').trim();
      var rate = REF.toNumber(vals[i][3]);
      if (!isFinite(rate)) rate = 0;

      if (tier === '1') out.r1 = rate;
      else if (tier === '2') out.r2 = rate;
      else if (tier === '3') out.r3 = rate;
      else if (tier === '3>') out.r3p = rate;
    }
    return out;
  };

  /* =========================
   *     Ц В Е Т А  К А Б И Н Е Т О В
   * ========================= */
  REF.readCabinetColorMap = function (platform /* 'OZON'|'WB'|'WILDBERRIES'|'OZ' */) {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.PARAMS);
    var map = new Map();
    if (!sh) return map;

    var last = sh.getLastRow();
    if (last < 2) return map;

    var want = String(platform||'').trim().toUpperCase();
    if (want === 'OZ') want = 'OZON';
    if (want === 'WB') want = 'WILDBERRIES';

    var names   = sh.getRange(2, 1, last - 1, 1).getDisplayValues(); // A
    var plats   = sh.getRange(2, 4, last - 1, 1).getDisplayValues(); // D
    var bgFills = sh.getRange(2, 7, last - 1, 1).getBackgrounds();   // G

    for (var i = 0; i < names.length; i++) {
      var cab  = String(names[i][0] || '').trim();
      var plat = String(plats[i][0] || '').trim().toUpperCase();
      var hex  = String(bgFills[i][0] || '').trim() || '#ffffff';
      if (!cab) continue;

      if (want) {
        var isOZ = (plat === 'OZON' || plat === 'OZ');
        var isWB = (plat === 'WILDBERRIES' || plat === 'WB');
        if ((want === 'OZON' && !isOZ) || (want === 'WILDBERRIES' && !isWB)) continue;
      }

      if (!map.has(cab)) map.set(cab, hex);
    }
    return map;
  };

  /* =========================
   *   К Л ю ч  и  С С  (legacy)
   * ========================= */
  REF.makeSSKey = function (cabinet, art) {
    return String(cabinet||'').trim() + '␟' + String(art||'').trim();
  };

  REF.readSSMap = function () {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.SS);
    var map = new Map();
    if (!sh) return map;

    var last = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (last < 2 || lastCol < 1) return map;

    var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    function findColIndex(name) {
      var target = String(name).toLowerCase();
      for (var i=0;i<headers.length;i++){
        var h = String(headers[i]||'').trim().toLowerCase();
        if (h === target) return i+1;
      }
      return -1;
    }

    var colCab = findColIndex('Кабинет');
    var colArt = findColIndex('Артикул');
    var colCC  = findColIndex('СС');
    if (colCab === -1 || colArt === -1 || colCC === -1) return map;

    var vals = sh.getRange(2, 1, last - 1, lastCol).getDisplayValues();
    for (var r=0;r<vals.length;r++){
      var row = vals[r];
      var cab = String(row[colCab-1]||'').trim();
      var art = String(row[colArt-1]||'').trim();
      if (!cab || !art) continue;

      var cc = REF.toNumber(row[colCC-1]);
      if (!isFinite(cc) || cc <= 0) continue;

      map.set(REF.makeSSKey(cab, art), cc);
    }
    return map;
  };

  /* =========================
   *      WB токены из «⚙️ Параметры»
   * ========================= */
  REF.buildWBTokenMapFromParams = function () {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.PARAMS);
    var map = new Map();
    if (!sh) return map;

    var last = sh.getLastRow();
    if (last < 2) return map;

    var rows = sh.getRange(2, 1, last - 1, 4).getDisplayValues();
    function dedupe(arr){
      var seen=new Set(), out=[];
      for (var i=0;i<arr.length;i++){
        var v=String(arr[i]||'').trim();
        if (!v || seen.has(v)) continue;
        seen.add(v); out.push(v);
      }
      return out;
    }

    for (var i=0;i<rows.length;i++){
      var cabinet = String(rows[i][0]==null?'':rows[i][0]).replace(/\u00A0/g,' ').trim();
      var type    = String(rows[i][1]||'').trim().toLowerCase();
      var token   = String(rows[i][2]||'').trim();
      var platUP  = String(rows[i][3]||'').trim().toUpperCase();
      if (!cabinet || !token) continue;
      if (!(platUP === 'WILDBERRIES' || platUP === 'WB')) continue;

      if (!map.has(cabinet)) map.set(cabinet, { prices:[], content:[], stats:[], supplies:[], any:[] });
      var rec = map.get(cabinet);
      rec.any.push(token);

      var parts = type.split(/[,\;/|]+/).map(function(s){return s.trim();});
      for (var p=0;p<parts.length;p++){
        var t = parts[p]; if (!t) continue;
        if (t.indexOf('цен')>-1 || t.indexOf('скид')>-1 || t.indexOf('аналит')>-1) rec.prices.push(token);
        if (t.indexOf('контент')>-1)  rec.content.push(token);
        if (t.indexOf('статист')>-1)  rec.stats.push(token);
        if (t.indexOf('постав')>-1)   rec.supplies.push(token);
      }

      rec.any      = dedupe(rec.any);
      rec.prices   = dedupe(rec.prices);
      rec.content  = dedupe(rec.content);
      rec.stats    = dedupe(rec.stats);
      rec.supplies = dedupe(rec.supplies);
    }
    return map;
  };

  REF.pickWBToken = function (cabinet, role, fallbackAny) {
    if (fallbackAny == null) fallbackAny = true;
    var map = REF.buildWBTokenMapFromParams();
    var key = String(cabinet==null?'':cabinet).replace(/\u00A0/g,' ').trim();
    var rec = map.get(key);
    if (!rec) return null;
    var pool = rec[role] || [];
    if (pool.length) return pool[0];
    return (fallbackAny && rec.any.length) ? rec.any[0] : null;
  };

  /* =========================
   *     Р А З Д Е Л Ы  (P:P)
   * ========================= */
  REF.readSectionPrefixes = function () {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.PARAMS);
    var out = [];
    if (!sh) return out;
    var last = sh.getLastRow();
    if (last < 2) return out;
    var vals = sh.getRange(2, 16, last - 1, 1).getDisplayValues(); // P
    for (var i=0;i<vals.length;i++){
      var v = String(vals[i][0]||'').trim().toLowerCase();
      if (v) out.push(v);
    }
    return out;
  };

  /* =========================
   *    «Т О В А Р»  и  «С С»
   * ========================= */

  // Преобразование артикульной записи в «товар» (как в 🍔СС!A)
  REF.toTovarFromArticle = function(platform, article) {
    var s = String(article||'').trim();
    if (s.length >= 3) s = s.substring(3); // часто артикулы в формате OZ_/WB_...
    s = s.replace(/_cat\d$/i, '');
    return s;
  };

  // «🍔 СС»!A:J → Map<tovar -> {cc, nal, vput, vpost, vpostOZ, vpostWB}>
  // ВНИМАНИЕ: vpost = vpostOZ + vpostWB (для обратной совместимости).
  REF.readSS_AJ_Map = function() {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.SS);
    var map = new Map();
    if (!sh) return map;

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 10) return map;

    var hdr = sh.getRange(1, 1, 1, 10).getDisplayValues()[0]; // A:J
    var norm = function(s){ return String(s||'').trim().toLowerCase(); };

    var idx = {
      tovar:   hdr.findIndex(function(h){ return norm(h) === 'товар'; }) + 1,
      ccud:    hdr.findIndex(function(h){ var n=norm(h); return n==='cc+упак+дост' || n==='сс+упак+дост' || n==='cc+упак+дост.'; }) + 1,
      nal:     hdr.findIndex(function(h){ return norm(h) === 'наличие'; }) + 1,
      vput:    hdr.findIndex(function(h){ return norm(h) === 'в пути'; }) + 1,
      vpostOZ: hdr.findIndex(function(h){ return norm(h) === 'в поставке oz'; }) + 1,
      vpostWB: hdr.findIndex(function(h){ return norm(h) === 'в поставке wb'; }) + 1
    };

    if (idx.tovar<=0 || idx.ccud<=0 || idx.nal<=0 || idx.vput<=0) return map;

    var hasOZ = idx.vpostOZ > 0;
    var hasWB = idx.vpostWB > 0;

    var vals = sh.getRange(2, 1, lastRow - 1, 10).getDisplayValues();
    for (var i=0;i<vals.length;i++){
      var row = vals[i];
      var key = String(row[idx.tovar-1]||'').trim();
      if (!key) continue;

      var cc     = REF.toNumber(row[idx.ccud - 1]);
      var nal    = REF.toNumber(row[idx.nal  - 1]);
      var vput   = REF.toNumber(row[idx.vput - 1]);
      var vpo    = hasOZ ? REF.toNumber(row[idx.vpostOZ - 1]) : 0;
      var vpw    = hasWB ? REF.toNumber(row[idx.vpostWB - 1]) : 0;

      cc   = (isFinite(cc)   ? cc   : 0);
      nal  = (isFinite(nal)  ? nal  : 0);
      vput = (isFinite(vput) ? vput : 0);
      vpo  = (isFinite(vpo)  ? vpo  : 0);
      vpw  = (isFinite(vpw)  ? vpw  : 0);

      map.set(key, {
        cc:    cc,
        nal:   nal,
        vput:  vput,
        vpost: vpo + vpw,
        vpostOZ: vpo,
        vpostWB: vpw
      });
    }
    return map;
  };

  // RichText для «Склад (M)»: [Наличие]+[В поставке] | [В пути]
  REF.buildWarehouseRich = function(nal, vpost, vput) {
    var n = Number(nal)||0, p = Number(vpost)||0, w = Number(vput)||0;
    var left=[], right=[];
    if (n>0) left.push({txt:String(n), color:'#000000'});
    if (p>0){ if (left.length) left.push({txt:'+', color:null}); left.push({txt:String(p), color:'#38761d'}); }
    if (w>0) right.push({txt:String(w), color:'#666666'});

    if (!left.length && !right.length) return '';

    var parts=[]; Array.prototype.push.apply(parts,left);
    if (left.length && right.length) parts.push({txt:' | ', color:null});
    Array.prototype.push.apply(parts,right);

    var text = parts.map(function(x){return x.txt;}).join('');
    var bold = SpreadsheetApp.newTextStyle().setBold(true).build();
    var b = SpreadsheetApp.newRichTextValue().setText(text);
    b.setTextStyle(bold);

    var cur=0;
    for (var i=0;i<parts.length;i++){
      var t = parts[i].txt, end = cur + t.length;
      if (parts[i].color) {
        var st = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor(parts[i].color).build();
        b.setTextStyle(cur, end, st);
      } else {
        b.setTextStyle(cur, end, bold);
      }
      cur = end;
    }
    return b.build();
  };

  // Отображение «⛓️ Параллель»
  REF.PARALLEL_DISPLAY = 'auto'; // 'auto' | 'article' | 'names'

  // База «Симка» из «🍔 СС» L:M
  REF.readSimkaBase = function () {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.SS);
    if (!sh) return 0;

    var last = sh.getLastRow();
    if (last < 2) return 0;

    var labels = sh.getRange(2, 12, last - 1, 1).getDisplayValues(); // L
    var vals   = sh.getRange(2, 13, last - 1, 1).getDisplayValues(); // M

    var base = 0;
    for (var i=0;i<labels.length;i++){
      var key = String(labels[i][0]||'').trim().toLowerCase();
      if (!key) continue;
      if (key === 'симка' || key.indexOf('симка') > -1) {
        var num = REF.toNumber(vals[i][0]);
        base = (isFinite(num) && num > 0) ? num : 0;
        if (base > 0) break;
      }
    }
    return base;
  };

  REF.isSimCardsCategory = function (ownCategory) {
    var s = String(ownCategory||'').replace(/\s+/g,' ').trim().toLowerCase();
    return s === 'симкарты';
  };

  /** Универсальный резолвер СС (используется в калькуляторе/выгрузках) */
  REF.resolveCCForArticle = function (platform, article, ownCategory, ssAJMap) {
    var map = ssAJMap || REF.readSS_AJ_Map();
    var tovar = REF.toTovarFromArticle(platform, article);
    var rec = map.get(tovar);
    var cc = rec && isFinite(Number(rec.cc)) ? Number(rec.cc) : 0;
    if (cc > 0) return cc;

    var simByCategory = REF.isSimCardsCategory(ownCategory);
    var simByPrefix   = String(tovar||'').toLowerCase().indexOf('sim0') === 0;
    if (simByCategory || simByPrefix) {
      var base = REF.readSimkaBase();
      var cc2x = (isFinite(base) && base > 0) ? base * 2 : 0;
      return cc2x > 0 ? cc2x : 0;
    }
    return 0;
  };

  REF.normCabinet = function (s) {
    return String(s == null ? '' : s)
      .replace(/[\u00A0\u2007\u202F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  };

  /* =========================
   *  ⚙️ Параметры: A:E + платформа
   * ========================= */

  // A..E: A=Кабинет, B=Тип токена/Роль, C=API KEY, D=Площадка, E=Краткое название
  REF.PARAMS_COLS = { CABINET:1, TOKEN_ROLE:2, API_KEY:3, PLATFORM:4, SHORT:5 };

  // Удобные ярлыки
  REF.PARAMS_RANGE_AE     = 'A:E';
  REF.PARAMS_SHORT_COL_A1 = 'E:E';

  // ⚠️ Платформа читается из ИМЕНОВАННОГО диапазона 'muff_mp'
  REF.PARAMS_PLATFORM_A1  = 'muff_mp';

  // Канонизация тэга платформы
  REF.platformCanon = function (p) {
    var raw = String(p||'').trim().toUpperCase();
    if (raw === 'OZON' || raw === 'OZ') return 'OZ';
    if (raw === 'WILDBERRIES' || raw === 'WB') return 'WB';
    return null;
  };

  // Универсальный геттер текущей платформы → 'OZ' | 'WB' | null
  REF.getCurrentPlatform = function () {
    var ss = SpreadsheetApp.getActive();
    var raw = '';
    try {
      raw = String(ss.getRangeByName(REF.PARAMS_PLATFORM_A1).getDisplayValue() || '').trim().toUpperCase();
    } catch (_) {
      try { raw = String(ss.getSheetByName(REF.SHEETS.PARAMS).getRange('I2').getDisplayValue() || '').trim().toUpperCase(); } catch(__){}
    }
    return REF.platformCanon(raw);
  };

  // Текущий выбранный кабинет из контрола (именованный диапазон muff_cabs)
  REF.getCabinetControlValue = function () {
    var ss = SpreadsheetApp.getActive();
    try {
      var rng = ss.getRangeByName(REF.CTRL_RANGE_A1);
      return String(rng && rng.getDisplayValue ? (rng.getDisplayValue() || '') : '').trim();
    } catch (_) { return ''; }
  };
  REF.getCabinetControlRange = function () {
    var ss = SpreadsheetApp.getActive();
    try { return ss.getRangeByName(REF.CTRL_RANGE_A1); } catch(_) { return null; }
  };

  /**
   * Map<Кабинет -> КраткоеНазвание>, с опц. фильтром по площадке
   */
  REF.readCabinetShortNameMap = function (platform) {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.PARAMS);
    var map = new Map();
    if (!sh) return map;

    var last = sh.getLastRow();
    if (last < 2) return map;

    var want = String(platform||'').trim().toUpperCase();
    if (want === 'OZ') want = 'OZON';
    if (want === 'WB') want = 'WILDBERRIES';

    var names = sh.getRange(2, REF.PARAMS_COLS.CABINET,  last - 1, 1).getDisplayValues(); // A
    var plats = sh.getRange(2, REF.PARAMS_COLS.PLATFORM, last - 1, 1).getDisplayValues(); // D
    var shorts= sh.getRange(2, REF.PARAMS_COLS.SHORT,    last - 1, 1).getDisplayValues(); // E

    for (var i=0;i<names.length;i++){
      var cabRaw   = String(names[i][0]  || '').trim();
      var platRaw  = String(plats[i][0]  || '').trim().toUpperCase();
      var shortRaw = String(shorts[i][0] || '').trim();
      if (!cabRaw) continue;

      if (want) {
        var isOZ = (platRaw === 'OZON' || platRaw === 'OZ');
        var isWB = (platRaw === 'WILDBERRIES' || platRaw === 'WB');
if ((want === 'OZON' && !isOZ) || (want === 'WILDBERRIES' && !isWB)) continue;

      }

      var key = REF.normCabinet(cabRaw);
      var val = shortRaw || cabRaw;
      if (!map.has(key)) map.set(key, val);
    }
    return map;
  };

  REF.getCabinetShortName = function (cabinet, platform) {
    var key = REF.normCabinet(cabinet);
    if (!key) return '';
    var map = REF.readCabinetShortNameMap(platform);
    return map.get(key) || key;
  };

  /* =========================
   *  Л О Г Г Е Р  S:U (⚙️ Параметры)
   * ========================= */
  // Колонки: S=19 (Обновления), T=20 (Время), U=21 (Кабинеты)
  // platformHint: 'OZON' | 'WILDBERRIES' | 'OZ' | 'WB' | undefined
  REF.logRun = function (opLabel, cabinetsArray, platformHint) {
    try {
      var ss = SpreadsheetApp.getActive();
      var sh = ss.getSheetByName(REF.SHEETS.PARAMS);
      if (!sh) return;

      var last = sh.getLastRow();
      if (last < 2) return;

      // Найти строку операции в S
      var ops = sh.getRange(2, 19, last - 1, 1).getDisplayValues(); // S
      var rowIndex = -1;
      for (var i = 0; i < ops.length; i++) {
        if (String(ops[i][0] || '').trim() === String(opLabel || '').trim()) { rowIndex = i + 2; break; }
      }
      if (rowIndex === -1) return;

      // Формат времени: dd.MM HH:mm (с ведущими нулями)
      var tz = ss.getSpreadsheetTimeZone() || 'Etc/GMT';
      var stamp = Utilities.formatDate(new Date(), tz, 'dd.MM HH:mm');

      // Если есть platformHint — нормализуем и используем как фильтр
      var want = String(platformHint || '').trim().toUpperCase();
      if (want === 'OZ') want = 'OZON';
      if (want === 'WB') want = 'WILDBERRIES';

      var seen = new Set(), out = [];
      var arr = Array.isArray(cabinetsArray) ? cabinetsArray : (cabinetsArray ? [cabinetsArray] : []);

      // Карта кратких имен с опциональным фильтром по площадке
      var shortMap = REF.readCabinetShortNameMap(want || undefined);

      for (var k = 0; k < arr.length; k++) {
        var full = REF.normCabinet(arr[k]);
        if (!full) continue;

        var shortName = shortMap.get(full);
        if (!shortName) {
          // Фолбэк: без фильтра (если в Params нет строки с такой площадкой)
          shortName = REF.getCabinetShortName(full, undefined) || full;
        }

        if (!seen.has(shortName)) {
          seen.add(shortName);
          out.push(shortName);
        }
      }

      sh.getRange(rowIndex, 20).setValue(stamp);           // T (Время)
      sh.getRange(rowIndex, 21).setValue(out.join(', '));  // U (Кабинеты)
    } catch (_){}
  };

  /* =========================
   *  Н А С Т Р О Й К И  (🎏 Форкаст!B:C)
   *  «Мин. запас на карточку» (ранее: «Минимальный запас»)
   *  B: FBS, шт   | C: <число>
   *  B: FBO, шт   | C: <число>
   *  B: FBS, дней | C: <число>
   *  B: FBO, дней | C: <число>
   * ========================= */
  REF.SETTINGS = REF.SETTINGS || {};
  REF.SETTINGS.FORECAST_SHEET = REF.SHEETS.FORECAST;
  REF.SETTINGS.MIN_STOCK_TITLES = ['Мин. запас на карточку', 'Минимальный запас']; // поддержка старого названия
  REF.SETTINGS.MIN_STOCK_LABELS = ['FBS, шт', 'FBO, шт', 'FBS, дней', 'FBO, дней'];

  REF.readMinStockPerCard = function () {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SETTINGS.FORECAST_SHEET);
    var out = { fbsUnits: 0, fboUnits: 0, fbsDays: 0, fboDays: 0 };
    if (!sh) return out;

    var lastRow = sh.getLastRow() || 1;

    // 1) Найти заголовок блока по merged-диапазонам B:C
    var titleRow = -1;
    try {
      var merged = sh.getMergedRanges() || [];
      for (var i = 0; i < merged.length; i++) {
        var r = merged[i];
        if (r.getColumn() !== 2) continue;           // начинается в B
        if (r.getNumColumns() < 2) continue;         // должен покрывать C
        var txt = String(r.getDisplayValue() || '').trim();
        if (!txt) continue;
        for (var t = 0; t < REF.SETTINGS.MIN_STOCK_TITLES.length; t++) {
          if (txt === REF.SETTINGS.MIN_STOCK_TITLES[t]) {
            titleRow = r.getRow();
            break;
          }
        }
        if (titleRow > 0) break;
      }
    } catch (_) {}

    // Фолбэк: если не нашли по merge, ищем точное совпадение текста в B
    if (titleRow <= 0) {
      var findTitles = REF.SETTINGS.MIN_STOCK_TITLES;
      var scanRows = Math.min(lastRow, 200);
      var colB = sh.getRange(1, 2, scanRows, 1).getDisplayValues(); // B1:Bscan
      outer:
      for (var r2 = 1; r2 <= scanRows; r2++) {
        var v = String(colB[r2 - 1][0] || '').trim();
        if (!v) continue;
        for (var tt = 0; tt < findTitles.length; tt++) {
          if (v === findTitles[tt]) { titleRow = r2; break outer; }
        }
      }
    }

    if (titleRow <= 0 || titleRow >= lastRow) return out;

    // 2) Собрать подряд непустые строки меток в B и значения в C
    var row = titleRow + 1;
    var options = []; // {label, valueRaw}
    while (row <= lastRow) {
      var label = String(sh.getRange(row, 2).getDisplayValue() || '').trim(); // B
      if (!label) break;
      var valRaw = sh.getRange(row, 3).getDisplayValue();                     // C
      options.push({ label: label, valueRaw: valRaw });
      row++;
    }

    if (!options.length) return out;

    function toNum(x) { return REF.toNumber ? REF.toNumber(x) : (parseFloat(String(x).replace(',', '.')) || 0); }

    for (var i2 = 0; i2 < options.length; i2++) {
      var key = options[i2].label;
      var val = toNum(options[i2].valueRaw);

      if (key === 'FBS, шт')   out.fbsUnits = isFinite(val) ? val : 0;
      else if (key === 'FBO, шт')   out.fboUnits = isFinite(val) ? val : 0;
      else if (key === 'FBS, дней') out.fbsDays  = isFinite(val) ? val : 0;
      else if (key === 'FBO, дней') out.fboDays  = isFinite(val) ? val : 0;
    }
    return out;
  };

  return REF;
})();
