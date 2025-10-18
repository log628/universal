/** =========================================================
 * Refs.gs — единый справочник и утилиты для проекта (OZ/WB)
 * ========================================================= */

(function(root){
  var REF = root.REF || (root.REF = {});

  /* ========= Л И С Т Ы ========= */
  REF.SHEETS = {
    ARTS_OZ:  '[OZ] Артикулы',
    ARTS_WB:  '[WB] Артикулы',
    FIZ_OZ:   '[OZ] Физ. оборот',
    FIZ_WB:   '[WB] Физ. оборот',
    PARAMS:   '⚙️ Параметры',
    RATES:    '🔖 Тарифы',
    SS:       '🍔 СС',
    CALC:     '⚖️ Калькулятор',
    FORECAST: '🎏 Форкаст'
  };

  /* ========= Named ranges ========= */
  REF.NAMED = REF.NAMED || {};
  REF.NAMED.MP_CTRL  = 'muff_mp';    // площадка OZ/WB
  REF.NAMED.CAB_CTRL = 'muff_cabs';  // выбор кабинета

  // Совместимость со старым кодом (некоторые места ждут A1-строку):
  // Ставим безопасный фолбэк — если вдруг где-то всё ещё дернут getRange(A1).
  REF.CTRL_RANGE_A1 = REF.CTRL_RANGE_A1 || 'B2';

  /* ========= У т и л и т ы ========= */
  REF.sheetName = function (key, fallback) {
    try { return (REF.SHEETS && REF.SHEETS[key]) || fallback || String(key||''); }
    catch(_){ return fallback || String(key||''); }
  };

  REF.toNumber = function (v) {
    if (v == null) return 0;
    if (typeof v === 'number' && isFinite(v)) return v;
    var s = String(v).trim();
    if (!s) return 0;
    s = s.replace(/\u00A0|\u2007|\u202F|\u2009/g, '').replace(/\s+/g, '');
    s = s.replace(/[^0-9.,\-]/g, '');
    if (s.indexOf(',') > -1 && s.indexOf('.') > -1) s = s.replace(/\./g,'').replace(',', '.');
    else if (s.indexOf(',') > -1) s = s.replace(',', '.');
    var n = parseFloat(s);
    return isFinite(n) ? n : 0;
  };
  REF.round2 = function(n){ var x=Number(n); return isFinite(x)?Math.round(x*100)/100:0; };
  REF.toComma = function(v){ return (v==null?'':String(v)).replace(/\./g, ','); };
  REF.toDot   = function(v){ return (v==null?'':String(v)).replace(/,/g, '.');  };

  REF.platformCanon = function (p) {
    var s = String(p||'').trim().toUpperCase();
    if (s==='OZ' || s==='OZON') return 'OZ';
    if (s==='WB' || s==='WILDBERRIES') return 'WB';
    return null;
  };

REF.getCurrentPlatform = function () {
  try {
    var ss = SpreadsheetApp.getActive();
    var r = ss.getRangeByName(REF.NAMED.MP_CTRL);
    var raw = r ? String(r.getDisplayValue() || '').trim() : '';
    var canon = REF.platformCanon(raw);
    return (canon === 'OZ' || canon === 'WB') ? canon : null;
  } catch (_) {
    return null;
  }
};


  REF.normCabinet = function (s) {
    return String(s == null ? '' : s)
      .replace(/[\u00A0\u2007\u202F]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  };

  REF.getCabinetControlRange = function () {
    try { return SpreadsheetApp.getActive().getRangeByName(REF.NAMED.CAB_CTRL); } catch(_){ return null; }
  };
  REF.getCabinetControlValue = function () {
    var r = REF.getCabinetControlRange();
    return r ? String(r.getDisplayValue()||'').trim() : '';
  };

  /* ========= Короткие имена кабинетов ========= */
  REF.readCabinetShortNameMap = function (platform /* optional: 'OZ'|'WB' */) {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.PARAMS);
    var map = new Map();
    if (!sh) return map;
    var last = sh.getLastRow(); if (last < 2) return map;

    var want = REF.platformCanon(platform);
    var A = sh.getRange(2,1,last-1,1).getDisplayValues(); // Кабинет
    var D = sh.getRange(2,4,last-1,1).getDisplayValues(); // МП
    var E = sh.getRange(2,5,last-1,1).getDisplayValues(); // Краткое
    for (var i=0;i<A.length;i++){
      var cab = String(A[i][0]||'').trim();
      var mp  = REF.platformCanon(D[i][0]);
      var sm  = String(E[i][0]||'').trim();
      if (!cab) continue;
      if (want && mp && mp !== want) continue;
      map.set(REF.normCabinet(cab), sm || cab);
    }
    return map;
  };
  REF.getCabinetShortName = function (cabinet, platform) {
    var m = REF.readCabinetShortNameMap(platform);
    return m.get(REF.normCabinet(cabinet)) || cabinet;
  };

  /* ========= Логгер (⚙️Параметры S:U) ========= */
  REF.logRun = function (opLabel, cabinetsArray, platformHint) {
    try {
      var ss = SpreadsheetApp.getActive();
      var sh = ss.getSheetByName(REF.SHEETS.PARAMS);
      if (!sh) return;
      var last = sh.getLastRow(); if (last < 2) return;

      var ops = sh.getRange(2,19,last-1,1).getDisplayValues(); // S
      var row = -1;
      for (var i=0;i<ops.length;i++){ if (String(ops[i][0]||'').trim() === String(opLabel||'').trim()) { row = i+2; break; } }
      if (row === -1) return;

      var tz = ss.getSpreadsheetTimeZone() || 'Etc/GMT';
      var stamp = Utilities.formatDate(new Date(), tz, 'dd.MM HH:mm');

      var want = REF.platformCanon(platformHint);
      var shortMap = REF.readCabinetShortNameMap(want);
      var seen = new Set(), out = [];
      var arr = Array.isArray(cabinetsArray) ? cabinetsArray : (cabinetsArray ? [cabinetsArray] : []);
      for (var k=0;k<arr.length;k++){
        var full = REF.normCabinet(arr[k]); if (!full) continue;
        var disp = shortMap.get(full) || arr[k];
        if (!seen.has(disp)) { seen.add(disp); out.push(disp); }
      }
      sh.getRange(row,20).setValue(stamp);      // T
      sh.getRange(row,21).setValue(out.join(', ')); // U
    } catch(_){}
  };

  /* ========= «СС» / «Симка» ========= */
  REF.toTovarFromArticle = function(platform, article) {
    var s = String(article||'').trim();
    if (s.length >= 3) s = s.substring(3); // OZ_/WB_
    s = s.replace(/_cat\d+$/i, '');
    return s;
  };
  REF.makeSSKey = function (cabinet, art) {
    return String(cabinet||'').trim() + '␟' + String(art||'').trim();
  };
REF.readSS_AJ_Map = function() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF.SHEETS.SS);
  var map = new Map();
  if (!sh) return map;

  var last = sh.getLastRow();
  var lastC = sh.getLastColumn();
  if (last < 2 || lastC < 2) return map;

  var hdr = sh.getRange(1,1,1,Math.min(10,lastC)).getDisplayValues()[0]
              .map(function(x){ return String(x||'').trim().toLowerCase(); });

  function find(nameArr){
    for (var i=0;i<hdr.length;i++) if (nameArr.indexOf(hdr[i])!==-1) return i+1;
    return 0;
  }

function findStartsWith(prefixArr){
  prefixArr = (prefixArr||[]).map(function(s){ return String(s||'').toLowerCase(); });
  for (var i=0;i<hdr.length;i++){
    var h = hdr[i];
    for (var j=0;j<prefixArr.length;j++){
      // startsWith без ES6
      if (h.lastIndexOf(prefixArr[j], 0) === 0) return i + 1;
    }
  }
  return 0;
}






  var iT   = find(['товар']);
  var iCC  = find(['cc+упак+дост','сс+упак+дост','cc+упак+дост.']);
var iNal = findStartsWith(['наличие']);   // опционально, "Наличие", "Наличие, шт", "Наличие (склад)" и т.п.
  var iPut = find(['в пути']);    // опционально
  // «В поставке» намеренно не ищем

  // Минимальный набор — Товар + СС
  if (!iT || !iCC) return map;

  var vals = sh.getRange(2,1,last-1,Math.min(10,lastC)).getDisplayValues();

  for (var r=0; r<vals.length; r++){
    var row = vals[r];
    var key = String(row[iT-1]||'').trim();
    if (!key) continue;

    var cc  = REF.toNumber(row[iCC-1]);
    var nal = iNal ? REF.toNumber(row[iNal-1]) : 0;
    var vput= iPut ? REF.toNumber(row[iPut-1]) : 0;

    map.set(key, {
      cc:    (isFinite(cc)   ? cc   : 0),
      nal:   (isFinite(nal)  ? nal  : 0),
      vput:  (isFinite(vput) ? vput : 0),
      vpost: 0 // больше не используем
    });
  }
  return map;
};

REF.readSimkaBase = function () {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF.SHEETS.SS);
  if (!sh) return 0;
  var last = sh.getLastRow(); if (last < 2) return 0;

  // Универсальный сканер пары столбцов: labels@colA, values@colB
  function scan(colA, colB) {
    var labels = sh.getRange(2, colA, last - 1, 1).getDisplayValues();
    var vals   = sh.getRange(2, colB, last - 1, 1).getDisplayValues();
    for (var i = 0; i < labels.length; i++) {
      var lab = String(labels[i][0] || '').trim().toLowerCase();
      if (!lab) continue;
      if (lab === 'симка' || lab.indexOf('симка') > -1) {
        var num = REF.toNumber(vals[i][0]);
        return (isFinite(num) && num > 0) ? num : 0;
      }
    }
    return 0;
  }

  // Основной источник теперь N:O (N=14, O=15)
  var base = scan(14, 15);
  if (base > 0) return base;

  // Бэкап на случай старых таблиц: L:M (L=12, M=13)
  var legacy = scan(12, 13);
  return legacy > 0 ? legacy : 0;
};

  REF.isSimCardsCategory = function (ownCategory) {
    var s = String(ownCategory||'').replace(/\s+/g,' ').trim().toLowerCase();
    return s === 'симкарты';
  };
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

  /* ========= Тарифы (если где-то понадобятся) ========= */
  REF.readTariffPercent = function (label) {
    var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(REF.SHEETS.RATES);
    if (!sh) return 0; var last = sh.getLastRow(); if (last<2) return 0;
    var AB = sh.getRange(2,1,last-1,2).getDisplayValues();
    for (var i=0;i<AB.length;i++){
      if (String(AB[i][0]||'').trim() === label) {
        var v = REF.toNumber(AB[i][1]); return isFinite(v)?v:0;
      }
    }
    return 0;
  };

  /* ========= Заголовки [OZ]/[WB] Артикулы ========= */
  REF.ARTS_HEADERS_BASE = [
    'Кабинет','Артикул','Отзывы','Рейтинг','Категория',
    'FBO','FBS','RFBS','Объем','Цена','SKU','Раздел','Своя категория'
  ];
  REF.getArtsHeaders = function (tag /* 'OZ'|'WB' */) {
    var hdr = REF.ARTS_HEADERS_BASE.slice();
    var t = String(tag||'').trim().toUpperCase();
    if (t) hdr[0] = '[ ' + t + ' ] Кабинет';
    return hdr;
  };
  REF.ARTS_COLS = { A:1,B:2,C:3,D:4,E:5,F:6,G:7,H:8,I:9,J:10,K:11,L:12,M:13 };
  REF.ARTS_TOTAL_COLS = 13;


  /* ========= Заголовки [OZ]/[WB] Физ. оборот ========= */
  REF.FIZ_HEADERS_BASE = [
    'Кабинет',
    'Артикул',
    'Остаток',
    'Скорость'
  ];
  REF.getFizHeaders = function (tag /* 'OZ'|'WB' */) {
    var hdr = REF.FIZ_HEADERS_BASE.slice();
    var t = String(tag||'').trim().toUpperCase();
    if (t === 'OZ') hdr[0] = '[ OZ ] Кабинет';
    else if (t === 'WB') hdr[0] = '[ WB ] Кабинет';
    return hdr;
  };
  REF.FIZ_COLS = { A:1,B:2,C:3,D:4 };
  REF.FIZ_TOTAL_COLS = 4;









  /* ========= Цвета кабинетов (по ЗАЛИВКЕ) =========
     platformTag: 'OZON'|'OZ'|'WILDBERRIES'|'WB'|null */
  REF.readCabinetColorMap = function(platformTag) {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.SHEETS.PARAMS);
    var map = new Map(); if (!sh) return map;
    var last = sh.getLastRow(); if (last < 2) return map;

    var hdr = sh.getRange(1,1,1,Math.max(7, sh.getLastColumn())).getDisplayValues()[0]
      .map(function(s){ return String(s||'').trim().toLowerCase(); });
    function idx(names){ for (var i=0;i<hdr.length;i++) if (names.indexOf(hdr[i])!==-1) return i+1; return 0; }

    var iCab   = idx(['кабинет','магазин','account','shop']) || 1; // A
    var iPlat  = idx(['мп','площадка','маркетплейс','platform']) || 4; // D
    var iColor = idx(['цвет_3','цвет','color']) || 7; // G — ЗАЛИВКА

    var A = sh.getRange(2,iCab,  last-1,1).getDisplayValues(); // имена
    var D = sh.getRange(2,iPlat, last-1,1).getDisplayValues(); // площадка
    var G = sh.getRange(2,iColor,last-1,1).getBackgrounds();   // HEX из заливки

    var want = REF.platformCanon(platformTag); // 'OZ'|'WB'|null
    for (var r=0;r<A.length;r++){
      var cab  = String(A[r][0]||'').trim(); if (!cab) continue;
      var plat = REF.platformCanon(D[r][0]);
      if (want && plat && plat !== want) continue;
      var hex  = String(G[r][0]||'').trim() || '#ffffff';
      if (!/^#?[0-9a-f]{6}$/i.test(hex)) hex = '#ffffff';
      if (hex[0] !== '#') hex = '#' + hex;
      if (hex.toLowerCase() === '#ffffff') continue; // белое — игнор
      map.set(cab, hex);
    }
    return map;
  };

  /* ========= «⚙️ Параметры» A:H ========= */
  REF.PARAMS = REF.PARAMS || {};
  REF.PARAMS.SHEET  = REF.SHEETS.PARAMS;
  REF.PARAMS.SCHEMA = { CAB:1, ID:2, KEY:3, MP:4, SHORT:5, TAX:6, COLOR:7, PICK:8 };
  REF.PARAMS.MP     = { OZON:'OZON', WB:'WILDBERRIES' };

  /** Карта токенов WB из A:H.
   * B «ID кабинета» — роли токена через запятую (например: "Контент, Статистика, Аналитика, Цены и скидки, Поставки")
   * C «API KEY»     — сам токен.
   * Возвращает Map(cabName -> {prices:[], content:[], stats:[], supplies:[], any:[]})
   */
  REF.buildWBTokenMapFromParams = function () {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.PARAMS.SHEET);
    var out = new Map(); if (!sh) return out;
    var last = sh.getLastRow(); if (last < 2) return out;
    var S = REF.PARAMS.SCHEMA;

    var vals = sh.getRange(2,1,last-1,S.PICK).getDisplayValues(); // A:H
    function toRoleKey(s){
      var t = String(s||'').toLowerCase().replace(/\s+/g,' ').trim();
      if (!t) return 'any';
      if (/контент/.test(t))        return 'content';
      if (/статист|аналит/.test(t)) return 'stats';
      if (/цен|скид/.test(t))       return 'prices';
      if (/постав/.test(t))         return 'supplies';
      if (/люба|any|общ/.test(t))   return 'any';
      return 'any';
    }
    function ensure(map, cab){
      if (!map.has(cab)) map.set(cab, { prices:[], content:[], stats:[], supplies:[], any:[] });
      return map.get(cab);
    }
    for (var i=0;i<vals.length;i++){
      var cab = String(vals[i][S.CAB-1]||'').trim();
      var id  = String(vals[i][S.ID -1]||'').trim(); // роли
      var key = String(vals[i][S.KEY-1]||'').trim(); // токен
      var mp  = String(vals[i][S.MP -1]||'').trim().toUpperCase();
      if (!cab || !key) continue;
      if (mp !== REF.PARAMS.MP.WB) continue;

      var roles = id ? id.split(',').map(function(s){return s.trim();}) : ['Любая'];
      var rec = ensure(out, cab);
      for (var r=0;r<roles.length;r++){ rec[toRoleKey(roles[r])].push(key); }
    }
    return out;
  };


/** Подбор токена WB по кабинету с доступом к "Цены и скидки".
 *  - Читает «⚙️ Параметры» A:H
 *  - Фильтр: D = Wildberries (любой регистр/вариант), A = cabinet
 *  - В B: роли через запятую, берём токен из C, если B содержит "цен" или "скид"
 *  - Порядок перечисления ролей не важен.
 *  - Возвращает первый подходящий токен, иначе null.
 */
REF.pickWBToken = function (cabinet, opts) {
  opts = opts || {};
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF.PARAMS.SHEET);
  if (!sh) return null;

  var last = sh.getLastRow();
  if (last < 2) return null;

  var S = REF.PARAMS.SCHEMA; // { CAB:1, ID:2, KEY:3, MP:4, SHORT:5, TAX:6, COLOR:7, PICK:8 }
  var vals = sh.getRange(2, 1, last - 1, S.PICK).getDisplayValues(); // A:H (display)
  var wantCab = REF.normCabinet(cabinet);

  var chosen = null;
  for (var i = 0; i < vals.length; i++) {
    var cab  = REF.normCabinet(vals[i][S.CAB - 1] || '');
    var roles= String(vals[i][S.ID  - 1] || '');
    var key  = String(vals[i][S.KEY - 1] || '').trim();
    var mp   = REF.platformCanon(vals[i][S.MP  - 1] || '');

    if (!key || cab !== wantCab || mp !== 'WB') continue;

    // Роли устойчиво парсим по вхождению "цен" или "скид" (без привязки к порядку слов)
    if (/цен|скид/i.test(roles)) {
      chosen = key;
      break; // берём первый подходящий
    }
  }
  return chosen;
};




  /** Отмеченные кабинеты по платформе (H=TRUE) */
  REF.readSelectedCabinets = function (mpTag /* 'OZON'|'WILDBERRIES' */) {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.PARAMS.SHEET);
    if (!sh) return { list: [], total: 0 };
    var last = sh.getLastRow(); if (last < 2) return { list: [], total: 0 };
    var S = REF.PARAMS.SCHEMA;
    var V = sh.getRange(2,1,last-1,S.PICK).getValues(); // RAW
    var list = [], total = 0, want = String(mpTag||'').toUpperCase();
    for (var i=0;i<V.length;i++){
      var cab = String(V[i][S.CAB-1]||'').trim();
      var mp  = String(V[i][S.MP -1]||'').trim().toUpperCase();
      var on  = !!V[i][S.PICK-1];
      if (!cab) continue;
      if (mp === want) { total++; if (on) list.push(cab); }
    }
    return { list:list, total:total };
  };

  /** Быстрые креды для OZON из A:H (D="OZON") */
  REF.readOZCredentials = function () {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(REF.PARAMS.SHEET);
    var arr = []; if (!sh) return arr;
    var last = sh.getLastRow(); if (last < 2) return arr;
    var S = REF.PARAMS.SCHEMA;
    var V = sh.getRange(2,1,last-1,S.PICK).getDisplayValues();
    for (var i=0;i<V.length;i++){
      var cab = String(V[i][S.CAB-1]||'').trim();
      var id  = String(V[i][S.ID -1]||'').trim();
      var key = String(V[i][S.KEY-1]||'').trim();
      var mp  = String(V[i][S.MP -1]||'').trim().toUpperCase();
      if (cab && id && key && mp === REF.PARAMS.MP.OZON) arr.push({ cab:cab, id:id, token:key });
    }
    return arr;
  };

})(this);
