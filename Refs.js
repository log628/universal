/** =========================================================
 * Refs.gs — справочники и утилиты (OZ/WB разделены)
 * Листы «[OZ] Артикулы» / «[WB] Артикулы» хранят A:M (13 колонок).
 * K=SKU, L=Раздел, M=Своя категория — базовые колонки листа (без «Наименование»).
 * Расширения дальше считаются «на лету» калькулятором.
 * ========================================================= */

var REF = (function () {
  var REF = {};

  /* =========================
   *        Л И С Т Ы
   * ========================= */
// ===== Листы (OZ/WB отдельно) =====
REF.SHEETS = {
  ARTS_OZ: '[OZ] Артикулы',
  ARTS_WB: '[WB] Артикулы',
  FIZ_OZ:  '[OZ] Физ. оборот',
  FIZ_WB:  '[WB] Физ. оборот',

  PARAMS:  '⚙️ Параметры',
  RATES:   '🔖 Тарифы',
  SS:      '🍔 СС',

  // ✅ Единое имя листа калькулятора
  CALC:    '⚖️ Калькулятор'
};

// ✅ Единый A1-диапазон контрола выбора кабинета
REF.CTRL_RANGE_A1 = 'B3:E4';

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

 // В REF (refs.gs) замените readCabinetColorMap на эту версию:
REF.readCabinetColorMap = function (platform /* 'OZON'|'WB'|'WILDBERRIES'|'OZ' */) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF.SHEETS.PARAMS);
  var map = new Map();
  if (!sh) return map;

  var last = sh.getLastRow();
  if (last < 2) return map;

  // Приводим площадку к каноническому виду
  var want = String(platform||'').trim().toUpperCase();
  if (want === 'OZ') want = 'OZON';
  if (want === 'WB') want = 'WILDBERRIES';

  // Читаем сразу нужные колонки: A (кабинет), D (площадка), G (цвет)
  var names   = sh.getRange(2, 1, last - 1, 1).getDisplayValues(); // A
  var plats   = sh.getRange(2, 4, last - 1, 1).getDisplayValues(); // D
  var bgFills = sh.getRange(2, 7, last - 1, 1).getBackgrounds();   // G

  for (var i = 0; i < names.length; i++) {
    var cab  = String(names[i][0] || '').trim();
    var plat = String(plats[i][0] || '').trim().toUpperCase();
    var hex  = String(bgFills[i][0] || '').trim() || '#ffffff';
    if (!cab) continue;

    // КЛЮЧЕВОЕ: сперва фильтр по площадке
    if (want) {
      var isOZ = (plat === 'OZON' || plat === 'OZ');
      var isWB = (plat === 'WILDBERRIES' || plat === 'WB');
      if ((want === 'OZON' && !isOZ) || (want === 'WILDBERRIES' && !isWB)) continue;
    }

    // Первое попадание выигрывает
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
  // A=Кабинет, B=Тип токена, C=API KEY, D=Площадка
  // Роли: prices | content | stats | supplies | any
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
  REF.toTovarFromArticle = function(platform, article) {
    var s = String(article||'').trim();
    if (s.length >= 3) s = s.substring(3);
    s = s.replace(/_cat\d$/i, '');
    return s;
  };

  // «🍔 СС»!A:J → Map<tovar -> {cc, nal, vput, vpost}>
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
      tovar: hdr.findIndex(function(h){ return norm(h) === 'товар'; }) + 1,
      ccud:  hdr.findIndex(function(h){ var n=norm(h); return n==='cc+упак+дост' || n==='сс+упак+дост'; }) + 1,
      nal:   hdr.findIndex(function(h){ return norm(h) === 'наличие'; }) + 1,
      vput:  hdr.findIndex(function(h){ return norm(h) === 'в пути'; }) + 1,
      vpost: hdr.findIndex(function(h){ return norm(h) === 'в поставке'; }) + 1
    };
    if (idx.tovar<=0 || idx.ccud<=0 || idx.nal<=0 || idx.vput<=0 || idx.vpost<=0) return map;

    var vals = sh.getRange(2, 1, lastRow - 1, 10).getDisplayValues();
    for (var i=0;i<vals.length;i++){
      var row = vals[i];
      var key = String(row[idx.tovar-1]||'').trim();
      if (!key) continue;

      var cc   = REF.toNumber(row[idx.ccud - 1]);
      var nal  = REF.toNumber(row[idx.nal  - 1]);
      var vput = REF.toNumber(row[idx.vput - 1]);
      var vpost= REF.toNumber(row[idx.vpost- 1]);

      map.set(key, {
        cc:   isFinite(cc)   ? cc   : 0,
        nal:  isFinite(nal)  ? nal  : 0,
        vput: isFinite(vput) ? vput : 0,
        vpost:isFinite(vpost)? vpost: 0
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

  /** Универсальный резолвер СС:
   *  1) tovar = article.slice(3) без "_cat<d>"
   *  2) cc = map[tovar].cc
   *  3) если cc<=0 и (ownCategory=="Симкарты" или tovar startsWith "sim0"):
   *       cc = readSimkaBase()*2
   */
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
    .replace(/[\u00A0\u2007\u202F]/g, ' ') // NBSP, figure space, narrow no-break
    .replace(/\s+/g, ' ')                  // схлопываем кратные пробелы
    .trim();
};
  return REF;
})();
