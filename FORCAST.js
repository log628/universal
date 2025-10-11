/** ======================================================================
 * FORCAST.gs — сборка листа «🎏 Форкаст» с учётом централизованного REF
 *  - Имена листов из REF.SHEETS
 *  - «🍔 СС» читается по заголовкам (A:J обязательно), L:M — курсы (опц.), O:P:Q — комплекты (опц.)
 *  - Флаг «Не закупается» из столбца с таким заголовком (A:J)
 *  - «В поставке» — единая колонка
 *  - Резолв товара из артикула делегирован в REF.toTovarFromArticle
 *  - Шапка на строке 2, данные с 3-й
 *  - Таблицы: E:I, K:O, Q:W
 *  - «Побелка» на E1:X[last] (фон, белые границы, сброс цвета шрифта и жирности)
 * ====================================================================== */

function buildForecast_All() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF && REF.SHEETS ? REF.SHEETS.FORECAST : '🎏 Форкаст') || ss.insertSheet('🎏 Форкаст');

  // ────────────────────────────────────────────────────────────────────────
  // У Т И Л И Т Ы
  // ────────────────────────────────────────────────────────────────────────
  function sheetName(key, fallback){ try{ return (REF && REF.SHEETS && REF.SHEETS[key]) || fallback; }catch(_){ return fallback; } }
  function num(v){ if (typeof REF!=='undefined' && REF.toNumber) return REF.toNumber(v);
    var s=String(v==null?'':v).replace(/\u00A0|\u2007|\u202F/g,'').replace(/\s+/g,'').replace(',', '.'); var n=parseFloat(s); return isFinite(n)?n:0; }
  function tovarFromArticle(platform, art){ try{ if (REF && REF.toTovarFromArticle) return REF.toTovarFromArticle(platform, art);}catch(_){}
    var s=String(art||'').trim(); if (s.length>=3) s=s.substring(3); return s.replace(/_cat\d+$/i,''); }
  function readNamedRaw(name){ try{ var rng=ss.getRangeByName(name); return rng ? rng.getDisplayValue() : ''; }catch(_){ return ''; } }
  function readNamedNumber(name, fallback){ var v=readNamedRaw(name); var n=num(v); return isFinite(n)?n:(isFinite(fallback)?fallback:0); }
  function norm(s){ return String(s==null?'':s).replace(/\u00A0|\u2007|\u202F/g,' ').replace(/\s+/g,' ').trim().toLowerCase(); }
  function isChecked(v){
    if (v === true) return true;
    var s = norm(v);
    return s === 'true' || s === '1' || s === 'x' || s === '✓' || s === 'да' || s === 'yes';
  }
  function isTrueNamed(name){
    var raw = readNamedRaw(name);
    if (raw === true) return true;
    return isChecked(raw);
  }
  function ceilToStep(x, step){
    var a = Math.max(0, Number(x)||0);
    var s = Math.max(1, Math.floor(Number(step)||0) || 1);
    return Math.ceil(a / s) * s;
  }
  function hdrIndex(hdrArr /* display row */, names /* string|string[] */){
    var hdr = (hdrArr||[]).map(function(v){return norm(v);});
    var list = Array.isArray(names)? names : [names];
    for (var i=0;i<hdr.length;i++){
      for (var j=0;j<list.length;j++){
        if (hdr[i] === norm(list[j])) return i+1;
      }
    }
    return 0;
  }

  // ────────────────────────────────────────────────────────────────────────
  // Ф И Л Ь Т Р Ы  (B:C)
  // ────────────────────────────────────────────────────────────────────────
  function findTitleRowInBC_(title){
    var want = norm(title);
    try{
      var merged = sh.getMergedRanges() || [];
      for (var i=0;i<merged.length;i++){
        var r = merged[i];
        var c0 = r.getColumn(), c1 = c0 + r.getNumColumns() - 1;
        if (c0 < 2 || c1 > 3) continue; // только в B:C
        var txt = norm(r.getDisplayValue()||'');
        if (txt === want) return r.getRow();
      }
    }catch(_){}
    var last = sh.getLastRow();
    var scan = Math.min(last, 500);
    var bc = sh.getRange(1,2,scan,2).getDisplayValues();
    for (var r=1;r<=scan;r++){
      if (norm(bc[r-1][0]) === want || norm(bc[r-1][1]) === want) return r;
    }
    return -1;
  }
  function readFilterList_BC_(title){
    var start = findTitleRowInBC_(title);
    var out = [];
    if (start <= 0) return out;
    var r = start + 1, last = sh.getLastRow();
    while (r <= last){
      var dispB = sh.getRange(r,2).getDisplayValue();
      if (dispB==='' || dispB==null) break;
      var rawB = sh.getRange(r,2).getValue();
      var rawC = (sh.getRange(r,3).getDisplayValue() || '').trim();
      out.push({ flag: isChecked(rawB), valueRaw: rawC, valueNorm: norm(rawC), row: r });
      r++;
    }
    return out;
  }

  // Параметры форкаста (кнопки)
  var TURNOVER   = readNamedNumber('forecast_button_turnover', 0);
  var EMPTYVAL   = readNamedNumber('forecast_button_empty',   0);
  var CHINAONLY  = isTrueNamed('forecast_button_chinaonly');
  var MINIMAL    = readNamedNumber('forecast_button_minimal', 0);
  var ROUNDSTEP  = readNamedNumber('forecast_button_round',   1);

  // Фильтры: площадки
  var platforms = readFilterList_BC_('Фильтр площадок');
  function isOZ(n){ return /^(ozon|oz)$/.test(n||''); }
  function isWB(n){ return /^(wildberries|wb)$/.test(n||''); }
  var wantOZ = platforms.some(function(x){ return x.flag && isOZ(x.valueNorm); });
  var wantWB = platforms.some(function(x){ return x.flag && isWB(x.valueNorm); });

  // Фильтры: категории
  var catRows = readFilterList_BC_('Фильтр категорий');
  var categoriesBlockFound = catRows.length > 0;
  var checkedCats = catRows.filter(function(x){ return x.flag; });
  var enabledCatOrderNorm = checkedCats.map(function(x){ return x.valueNorm; });
  var catSet = new Set(enabledCatOrderNorm);
  var mustBlockByCats = (categoriesBlockFound && enabledCatOrderNorm.length === 0);

  // ────────────────────────────────────────────────────────────────────────
  // И М Е Н А  Л И С Т О В
  // ────────────────────────────────────────────────────────────────────────
  var SH_ARTS_OZ = sheetName('ARTS_OZ', '[OZ] Артикулы');
  var SH_ARTS_WB = sheetName('ARTS_WB', '[WB] Артикулы');
  var SH_FIZ_OZ  = sheetName('FIZ_OZ',  '[OZ] Физ. оборот');
  var SH_FIZ_WB  = sheetName('FIZ_WB',  '[WB] Физ. оборот');
  var SH_SS      = sheetName('SS',      '🍔 СС');

  // ────────────────────────────────────────────────────────────────────────
  // Ч Т Е Н И Е  «🍔 СС»
  // ────────────────────────────────────────────────────────────────────────
  function readSS_All_(){
    var s = ss.getSheetByName(SH_SS);
    var out = {
      goods: new Map(), // tovar -> { brand, model, ccCur, currency, nal, vput, vpostSum, notBuy }
      kits:  [],        // [{kit, comp, coef}]
      rates: new Map(), // norm(currency) -> rate
      notBuySet: new Set()
    };
    if (!s) return out;
    var lr = s.getLastRow(); if (lr < 2) return out;

    var lc = s.getLastColumn();
    var hdr = s.getRange(1,1,1,lc).getDisplayValues()[0];

    // Индексы по заголовкам (A:J обязательно, остальное — опционально)
    var cTovar = hdrIndex(hdr, 'товар');
    var cBrand = hdrIndex(hdr, 'производитель');
    var cModel = hdrIndex(hdr, 'модель');
    var cCCcur = hdrIndex(hdr, ['cc в валюте', 'сс в валюте']);
    var cCurr  = hdrIndex(hdr, 'валюта');
    var cCCUD  = hdrIndex(hdr, ['cc+упак+дост','сс+упак+дост','cc+упак+дост.']);
    var cNal   = hdrIndex(hdr, 'наличие');
    var cVput  = hdrIndex(hdr, 'в пути');
    var cVpost = hdrIndex(hdr, 'в поставке');
    var cOff   = hdrIndex(hdr, 'не закупается');

    var readCols = Math.min(lc, Math.max(cTovar,cBrand,cModel,cCCcur,cCurr,cCCUD,cNal,cVput,cVpost,cOff,10));
    var rowsAJ = s.getRange(2,1,lr-1,readCols).getDisplayValues();

    for (var i=0;i<rowsAJ.length;i++){
      var row = rowsAJ[i];
      var tv = String(row[(cTovar||1)-1]||'').trim(); if (!tv) continue;

      var brand = cBrand ? String(row[cBrand-1]||'').trim() : '';
      var model = cModel ? String(row[cModel-1]||'').trim() : '';
      var ccCur = cCCcur? num(row[cCCcur-1]) : 0;
      var curr  = cCurr ? String(row[cCurr-1]||'').trim() : '';
      var nal   = cNal   ? num(row[cNal-1])   : 0;
      var vput  = cVput  ? num(row[cVput-1])  : 0;
      var vpost = cVpost ? num(row[cVpost-1]) : 0;

      var notBuy = cOff ? (norm(row[cOff-1]) === 'да') : false;
      if (notBuy) out.notBuySet.add(tv);

      out.goods.set(tv, {
        brand: brand,
        model: model,
        ccCur: isFinite(ccCur)?ccCur:0,
        currency: curr,
        nal: isFinite(nal)?nal:0,
        vput: isFinite(vput)?vput:0,
        vpostSum: isFinite(vpost)?vpost:0,
        notBuy: notBuy
      });
    }

    // Курсы L:M — опционально
    if (lc >= 13){
      var labels = s.getRange(2,12,lr-1,1).getDisplayValues(); // L
      var rates  = s.getRange(2,13,lr-1,1).getDisplayValues(); // M
      for (var r=0;r<labels.length;r++){
        var name = String(labels[r][0]||'').trim();
        if (!name) continue;
        var rate = num(rates[r][0]);
        out.rates.set(norm(name), isFinite(rate)?rate:0);
      }
    }

    // Комплекты O:P:Q — опционально
    if (lc >= 17){
      var kits = s.getRange(2,15,lr-1,3).getDisplayValues(); // O:P:Q
      for (var k=0;k<kits.length;k++){
        var kit  = String(kits[k][0]||'').trim();
        var comp = String(kits[k][1]||'').trim();
        var coef = num(kits[k][2]);
        if (!kit || !comp) continue;
        var c = isFinite(coef)?coef:0;
        if (c <= 0) continue;
        out.kits.push({ kit: kit, comp: comp, coef: c });
      }
    }

    return out;
  }
  var SS = readSS_All_();

  function isRubleCurrency_(s){
    var x = norm(s);
    if (!x) return false;
    if (x === 'рубль' || x === 'ruble') return true;
    if (x === 'руб' || x === 'rub' || x === 'rur') return true;
    if (x.indexOf('руб') === 0) return true;
    return false;
  }

  // ────────────────────────────────────────────────────────────────────────
  // Ч Т Е Н И Е  А Р Т И К У Л О В  + С В О Я  К А Т Е Г О Р И Я
  // ────────────────────────────────────────────────────────────────────────
  function readArticlesWithCategory_(sheetName_){
    var s = ss.getSheetByName(sheetName_);
    var out = [];
    if (!s) return out;
    var lr = s.getLastRow(), lc = s.getLastColumn();
    if (lr < 2 || lc < 2) return out;
    var hdrNorm = s.getRange(1,1,1,lc).getDisplayValues()[0].map(norm);
    var idxArt = hdrNorm.indexOf('артикул') + 1; if (idxArt <= 0) idxArt = 2;
    var idxOwn = hdrNorm.indexOf('своя категория') + 1; if (idxOwn <= 0) idxOwn = 13;

    var vals = s.getRange(2,1,lr-1,lc).getDisplayValues();
    for (var i=0;i<vals.length;i++){
      var art = String(vals[i][idxArt-1]||'').trim(); if (!art) continue;
      var catRaw = vals[i][idxOwn-1] || '';
      var catN = norm(catRaw);
      if (mustBlockByCats) continue;
      if (categoriesBlockFound && catSet.size > 0) { if (!catSet.has(catN)) continue; }
      out.push({ art: art, catNorm: catN });
    }
    return out;
  }
  var artsOZ = wantOZ ? readArticlesWithCategory_(SH_ARTS_OZ) : [];
  var artsWB = wantWB ? readArticlesWithCategory_(SH_ARTS_WB) : [];

  // ────────────────────────────────────────────────────────────────────────
  // Г Р У П П И Р О В К А  П О  Т О В А Р А М  (+ «не закупается»)
  // ────────────────────────────────────────────────────────────────────────
  var byTovar = new Map();
  var tovarCats = new Map();
  function add(platformTag, rec){
    var tv = tovarFromArticle(platformTag, rec.art);
    if (!tv) return;
    if (SS.notBuySet.has(tv)) return; // исключаем «не закупается»

    if (!byTovar.has(tv)) byTovar.set(tv, { oz:new Set(), wb:new Set() });
    byTovar.get(tv)[platformTag==='OZ'?'oz':'wb'].add(rec.art);

    if (!tovarCats.has(tv)) tovarCats.set(tv, new Set());
    if (rec.catNorm) tovarCats.get(tv).add(rec.catNorm);
  }
  artsOZ.forEach(function(r){ add('OZ', r); });
  artsWB.forEach(function(r){ add('WB', r); });

  function categoryIndexForTovar(tv){
    if (!enabledCatOrderNorm.length) return 1e9;
    var cats = tovarCats.get(tv);
    if (!cats || !cats.size) return 1e9;
    var best = 1e9;
    enabledCatOrderNorm.forEach(function(cat, idx){
      if (cats.has(cat) && idx < best) best = idx;
    });
    return best;
  }

  var TOVARS_ALL = Array.from(byTovar.keys()).sort(function(a,b){
    var ia = categoryIndexForTovar(a), ib = categoryIndexForTovar(b);
    if (ia !== ib) return ia - ib;
    return a.localeCompare(b);
  });

  var TOVARS_CHINA = TOVARS_ALL.filter(function(tv){
    if (!CHINAONLY) return true;
    var rec = SS.goods.get(tv);
    if (!rec) return false;
    if (!rec.currency) return false;
    return !isRubleCurrency_(rec.currency);
  });

  // ────────────────────────────────────────────────────────────────────────
  // Ф И З . О Б О Р О Т
  // ────────────────────────────────────────────────────────────────────────
  function readFizMapsByArticle_(sheetName_, enabled){
    var maps = { fbo:new Map(), spd:new Map() };
    if (!enabled) return maps;
    var s = ss.getSheetByName(sheetName_);
    if (!s) return maps;
    var lr = s.getLastRow(); if (lr < 2 || s.getLastColumn() < 4) return maps;
    var vals = s.getRange(2,2,lr-1,3).getDisplayValues(); // B=art, C=fbo, D=spd
    for (var i=0;i<vals.length;i++){
      var art = String(vals[i][0]||'').trim(); if (!art) continue;
      var fbo = num(vals[i][1]) || 0;
      var spd = num(vals[i][2]) || 0;
      maps.fbo.set(art, (maps.fbo.get(art)||0) + fbo);
      maps.spd.set(art, (maps.spd.get(art)||0) + spd);
    }
    return maps;
  }
  var fizOZ = readFizMapsByArticle_(SH_FIZ_OZ, wantOZ);
  var fizWB = readFizMapsByArticle_(SH_FIZ_WB, wantWB);

  // ────────────────────────────────────────────────────────────────────────
  // Р А С Ч Ё Т  «П О Т Р Е Б Н О С Т И»  Д Л Я  А Р Т И К У Л О В  (Q:W)
  // ────────────────────────────────────────────────────────────────────────
  function calcNeedAndTag(FBO_OZ,FBO_WB,SPD_OZ,SPD_WB){
    var bothZero = (FBO_OZ===0 && FBO_WB===0 && SPD_OZ===0 && SPD_WB===0);
    if (bothZero) return { val: EMPTYVAL, tag: 'orange' };
    var needOZ = Math.max(SPD_OZ*TURNOVER - FBO_OZ, 0);
    var needWB = Math.max(SPD_WB*TURNOVER - FBO_WB, 0);
    var v = Math.floor(needOZ + needWB);
    if (v > 0) return { val: v, tag: 'blue' };
    return { val: 0, tag: null };
  }

  // ────────────────────────────────────────────────────────────────────────
  // П О Д Г О Т О В К А  Д А Н Н Ы Х  Q:W
  // ────────────────────────────────────────────────────────────────────────
  var rowsTZ = [];
  var productIdx = [];
  var sumNeedByTovar_raw = new Map();
  var nothingToCollect = (!wantOZ && !wantWB) || (categoriesBlockFound && enabledCatOrderNorm.length===0);

  var T_sorted = Array.from(TOVARS_CHINA).sort(function(a,b){
    var ia = categoryIndexForTovar(a), ib = categoryIndexForTovar(b);
    if (ia!==ib) return ia-ib;
    return a.localeCompare(b);
  });

  if (!nothingToCollect){
    for (var t=0;t<T_sorted.length;t++){
      var tv = T_sorted[t];
      var bucket = byTovar.get(tv) || { oz:new Set(), wb:new Set() };
      var ozList = Array.from(bucket.oz).sort((a,b)=>a.localeCompare(b));
      var wbList = Array.from(bucket.wb).sort((a,b)=>a.localeCompare(b));

      var ozCount=ozList.length, wbCount=wbList.length;
      var sumFBO_OZ=0,sumFBO_WB=0,sumSPD_OZ=0,sumSPD_WB=0,sumNeed=0;

      var idx = rowsTZ.length;
      rowsTZ.push(['  '+tv, (ozCount+' | '+wbCount), 0,0,0,0,0]); // Q..W
      productIdx.push(idx);

      function pushArts(tag, list){
        for (var i=0;i<list.length;i++){
          var art = list[i];
          var FBO_OZ = (tag==='oz') ? (fizOZ.fbo.get(art)||0) : 0;
          var FBO_WB = (tag==='wb') ? (fizWB.fbo.get(art)||0) : 0;
          var SPD_OZ = (tag==='oz') ? (fizOZ.spd.get(art)||0) : 0;
          var SPD_WB = (tag==='wb') ? (fizWB.spd.get(art)||0) : 0;

          var q = calcNeedAndTag(FBO_OZ,FBO_WB,SPD_OZ,SPD_WB);

          sumFBO_OZ+=FBO_OZ; sumFBO_WB+=FBO_WB; sumSPD_OZ+=SPD_OZ; sumSPD_WB+=SPD_WB; sumNeed+=q.val;

          rowsTZ.push(['       '+art, (tag==='oz'?'oz':'wb'), FBO_OZ||0, FBO_WB||0, SPD_OZ||0, SPD_WB||0, q.val||0]);
        }
      }
      if (wantOZ) pushArts('oz', ozList);
      if (wantWB) pushArts('wb', wbList);

      rowsTZ[idx][2]=sumFBO_OZ; rowsTZ[idx][3]=sumFBO_WB;
      rowsTZ[idx][4]=sumSPD_OZ; rowsTZ[idx][5]=sumSPD_WB;
      rowsTZ[idx][6]=sumNeed;

      sumNeedByTovar_raw.set(tv, sumNeed);
    }
  }

  // ────────────────────────────────────────────────────────────────────────
  // К О М П Л Е К Т Ы
  // ────────────────────────────────────────────────────────────────────────
  var compSecondType = new Set(); // товары с O==P
  (SS.kits||[]).forEach(function(edge){
    if (edge.kit && edge.comp && edge.kit === edge.comp) compSecondType.add(edge.kit);
  });

  function isBrandKit(tv){
    var rec = SS.goods.get(tv);
    if (!rec) return false;
    return norm(rec.brand) === 'комплект';
  }

  var selfCoef = new Map(); // tv -> sum Q для строк O==P==tv
  (SS.kits||[]).forEach(function(edge){
    if (edge.kit && edge.comp && edge.kit === edge.comp){
      selfCoef.set(edge.kit, (selfCoef.get(edge.kit)||0) + edge.coef);
    }
  });

  var addFromKits = new Map();
  (SS.kits||[]).forEach(function(edge){
    var kit = edge.kit, comp = edge.comp, c = edge.coef;
    if (!kit || !comp || c<=0) return;
    if (kit === comp) return; // self — в base для второго типа

    if (!sumNeedByTovar_raw.has(kit)) return;
    var zKit = sumNeedByTovar_raw.get(kit) || 0;
    if (zKit <= 0) return;

    addFromKits.set(comp, (addFromKits.get(comp)||0) + zKit * c);
  });

  // ────────────────────────────────────────────────────────────────────────
  // Т А Б Л И Ц А  K:O — по товарам (без обычных комплектов)
  // ────────────────────────────────────────────────────────────────────────
  var rowsNR = [];
  if (!nothingToCollect){
    var listForNR = T_sorted.filter(function(tv){
      return !isBrandKit(tv);
    });

    for (var i=0;i<listForNR.length;i++){
      var tv = listForNR[i];
      var ssrec = (SS.goods.get(tv) || { nal:0, vput:0, vpostSum:0 });

      var zRaw = sumNeedByTovar_raw.get(tv) || 0;
      var base;
      if (compSecondType.has(tv)){
        var sc = selfCoef.get(tv) || 0;
        base = zRaw * sc;
      } else {
        base = zRaw;
      }

      var plusFromKits = addFromKits.get(tv) || 0;
      var P_total = base + plusFromKits;

      var P_disp = '';
      if (base > 0 && plusFromKits > 0) P_disp = String(base) + '+' + String(plusFromKits);
      else if (base <= 0 && plusFromKits > 0) P_disp = '+' + String(plusFromKits);
      else if (base > 0 && plusFromKits <= 0) P_disp = String(base);
      else P_disp = '';

      var baseKup = Math.max(0, (P_total + MINIMAL) - (ssrec.nal + ssrec.vput));
      var kup  = (baseKup < 3) ? 0 : ceilToStep(baseKup, ROUNDSTEP);

      rowsNR.push(['  '+tv, kup, P_disp, ssrec.nal||0, ssrec.vput||0]);
    }
  }

  // ────────────────────────────────────────────────────────────────────────
  // Т А Б Л И Ц А  E:I — «к закупу» из K:O (qty>0)
  // ────────────────────────────────────────────────────────────────────────
  function catKeyForTovar(tv){
    var idx = categoryIndexForTovar(tv);
    var name = '';
    if (idx !== 1e9){
      var cats = tovarCats.get(tv);
      for (var i=0;i<enabledCatOrderNorm.length;i++){
        if (cats && cats.has(enabledCatOrderNorm[i])) { name = enabledCatOrderNorm[i]; break; }
      }
    }
    return { idx: idx, name: name };
  }

  var rowsFJData = [];
  if (rowsNR.length){
    for (var i=0;i<rowsNR.length;i++){
      var tvDisp = String(rowsNR[i][0]||'');
      var tv = tvDisp.startsWith('  ') ? tvDisp.substring(2) : tvDisp;
      var qty = Number(rowsNR[i][1])||0;
      if (qty > 0){
        var rec = SS.goods.get(tv) || {};
        var brand = rec.brand || '';
        var model = rec.model || '';
        var ccCur = Number(rec.ccCur)||0;
        var curName = rec.currency || '';
        var rate = SS.rates.get(norm(curName)) || 0;
        var cost = (ccCur * rate) * qty;

        var ck = catKeyForTovar(tv);
        rowsFJData.push({
          brand: brand,
          catIdx: ck.idx,
          catName: ck.name,
          tv: tv,
          qty: qty,
          cost: cost,
          model: model
        });
      }
    }
  }
  rowsFJData.sort(function(a,b){
    var bb = a.brand.localeCompare(b.brand);
    if (bb !== 0) return bb;
    if (a.catIdx !== b.catIdx) return a.catIdx - b.catIdx;
    return a.tv.localeCompare(b.tv);
  });
  var rowsFJ = rowsFJData.map(function(r){
    return ['  '+r.tv, '  '+r.brand, '  '+r.model, r.qty, r.cost];
  });

  // ────────────────────────────────────────────────────────────────────────
  // Р И С О В К А
  // ────────────────────────────────────────────────────────────────────────
  var WHITE = '#ffffff';
  var BLACK = '#000000';
  var SOLID = SpreadsheetApp.BorderStyle.SOLID;
  var HDR_ROW  = 2; // строка заголовков
  var DATA_ROW = 3; // старт данных

  // 0) Общая «побелка» по прямоугольнику E1:X[last]:
  var lastRowSheet = Math.max(sh.getMaxRows(), HDR_ROW);
  var bleach = sh.getRange(1,5, lastRowSheet, 24-5+1); // E..X
  bleach.clearContent()
        .setBackground(WHITE)
        .setBorder(true,true,true,true,true,true,WHITE,SOLID) // белая «сетка»
        .setFontFamily('Roboto').setFontSize(10)
        .setFontColor('#000000').setFontWeight('normal');

  // Базовые ширины колонок по заданию
  sh.setColumnWidth(1, 25);   // A
  sh.setColumnWidth(2, 60);   // B
  sh.setColumnWidth(3,160);   // C
  sh.setColumnWidth(4, 50);   // D
  sh.setColumnWidth(10,35);   // J (прокладка)
  sh.setColumnWidth(16,35);   // P (прокладка)

  // Вспомогательный локальный «сброс границ» (шапка+данные области вставки)
  function clearBordersLocal_(row1, col1, height, width){
    sh.getRange(row1, col1, Math.max(height,1), Math.max(width,1))
      .setBorder(false,false,false,false,false,false,null,null);
  }

  // ── Таблица Q:W ─────────────────────────────────────────────────────────
  (function draw_QW(){
    var totalData = Math.max(rowsTZ.length, 0);
    var usedRow = totalData ? (DATA_ROW - 1 + rowsTZ.length) : HDR_ROW;

    // Локальный «сброс границ» (шапка+данные)
    clearBordersLocal_(HDR_ROW, 17, usedRow - HDR_ROW + 1, 7); // Q..W

    // Шапка (строка 2)
    sh.getRange(HDR_ROW,17).setValue('Товар / Артикул'); // Q2
    sh.getRange(HDR_ROW,18).setValue('МП');               // R2
    sh.getRange(HDR_ROW,19,1,1).setValue('Остаток FBO');  // S2:T2 merged
    sh.getRange(HDR_ROW,21,1,1).setValue('Скорость');     // U2:V2 merged
    sh.getRange(HDR_ROW,23).setValue('Потребность');      // W2
    sh.getRange(HDR_ROW,19,1,2).merge(); // S2:T2
    sh.getRange(HDR_ROW,21,1,2).merge(); // U2:V2

    var DARK='#434343', FBO_D='#274e13', SPD_D='#1c4587', NEED_D='#741b47';
    sh.getRange(HDR_ROW,17,1,2).setBackground(DARK);     // Q:R
    sh.getRange(HDR_ROW,19,1,2).setBackground(FBO_D);    // S:T
    sh.getRange(HDR_ROW,21,1,2).setBackground(SPD_D);    // U:V
    sh.getRange(HDR_ROW,23,1,1).setBackground(NEED_D);   // W

    var header = sh.getRange(HDR_ROW,17,1,7);
    header.setFontColor('#ffffff').setFontFamily('Roboto').setFontSize(10)
          .setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setFontWeight('bold')
          .setBorder(true,true,true,true,true,true);

    // Данные (с 3-й строки)
    if (rowsTZ.length) sh.getRange(DATA_ROW,17,rowsTZ.length,7).setValues(rowsTZ);
    var totalRows = Math.max(rowsTZ.length, 1);
    sh.getRange(DATA_ROW,17,totalRows,7).setFontFamily('Roboto').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(DATA_ROW,17,totalRows,1)
      .setHorizontalAlignment('left').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    if (totalRows>0){
      sh.getRange(DATA_ROW,19,totalRows,2).setBackground('#d9ead3'); // S:T
      sh.getRange(DATA_ROW,21,totalRows,2).setBackground('#c9daf8'); // U:V
      sh.getRange(DATA_ROW,23,totalRows,1).setBackground('#ead1dc'); // W
    }

    var INT='#,##0;-#,##0;;@', DEC2='#,##0.00;-#,##0.00;;@';
    if (totalRows>0){
      sh.getRange(DATA_ROW,19,totalRows,2).setNumberFormat(INT);  // S:T
      sh.getRange(DATA_ROW,21,totalRows,2).setNumberFormat(DEC2); // U:V
      sh.getRange(DATA_ROW,23,totalRows,1).setNumberFormat(INT);  // W
    }

    // Сброс стиля W (потребность) перед точечной раскраской
    if (totalRows>0){
      sh.getRange(DATA_ROW,23,totalRows,1).setFontColor('#000000').setFontWeight('normal');
    }

    var prodExcelRows = new Set(productIdx.map(function(i){ return DATA_ROW + i; }));

    for (var i=0;i<rowsTZ.length;i++){
      var r = DATA_ROW + i;
      if (prodExcelRows.has(r)){
        // строка ТОВАРА
        sh.getRange(r,17,1,7).setBackground('#cccccc').setFontWeight('bold')
          .setBorder(true,null,true,null,null,null,BLACK,SOLID);
        sh.getRange(r,23,1,1).setFontColor('#000000');
      } else {
        // строка АРТИКУЛА → Q:R светло-серый фон
        sh.getRange(r,17,1,2).setBackground('#f3f3f3'); // Q:R
        var row = rowsTZ[i];
        var FBO_OZ=Number(row[2]), FBO_WB=Number(row[3]), SPD_OZ=Number(row[4]), SPD_WB=Number(row[5]), NEED=Number(row[6]);
        if (FBO_OZ===0 && FBO_WB===0 && SPD_OZ===0 && SPD_WB===0) {
          sh.getRange(r,23).setFontWeight('bold').setFontColor('#e69138');
        } else if (NEED>0) {
          sh.getRange(r,23).setFontWeight('bold').setFontColor('#0000ff');
        }
      }
    }

    // Контуры и вертикальные «грани»
    function rb(c){ sh.getRange(HDR_ROW,c,Math.max(usedRow-HDR_ROW+1,1),1).setBorder(null,null,null,true,null,null,BLACK,SOLID); }
    function lb(c){ sh.getRange(HDR_ROW,c,Math.max(usedRow-HDR_ROW+1,1),1).setBorder(null,true,null,null,null,null,BLACK,SOLID); }
    lb(17); rb(23); // внешний контур Q..W
    rb(18);         // R | S
    rb(20);         // T | U
    rb(22);         // V | W
    if (rowsTZ.length) sh.getRange(usedRow,17,1,7).setBorder(null,null,true,null,null,null,BLACK,SOLID);

    // Ширины
    sh.autoResizeColumn(17); sh.setColumnWidth(17, sh.getColumnWidth(17)+50); // Q
    sh.setColumnWidth(18, 55); // R = 55
    sh.setColumnWidths(19, 2, 65); // S:T = 65
    sh.setColumnWidths(21, 2, 65); // U:V = 65
    sh.setColumnWidth(23,110); // W
  })();

  // ── Таблица K:O ─────────────────────────────────────────────────────────
  (function draw_KO(){
    var totalData = Math.max(rowsNR.length, 0);
    var usedRow = totalData ? (DATA_ROW - 1 + rowsNR.length) : HDR_ROW;

    // Локальный «сброс границ»
    clearBordersLocal_(HDR_ROW, 11, usedRow - HDR_ROW + 1, 5); // K..O

    // Шапка
    sh.getRange(HDR_ROW,11).setValue('Товар');       // K2
    sh.getRange(HDR_ROW,12).setValue('К закупу');    // L2
    sh.getRange(HDR_ROW,13).setValue('Потребность'); // M2 (строковое "base+add")
    sh.getRange(HDR_ROW,14).setValue('Налич');       // N2
    sh.getRange(HDR_ROW,15).setValue('Путь');        // O2

    var DARK='#434343', NEED_D='#741b47', WH_D='#783f04';
    sh.getRange(HDR_ROW,11,1,2).setBackground(DARK);     // K:L
    sh.getRange(HDR_ROW,13,1,1).setBackground(NEED_D);   // M
    sh.getRange(HDR_ROW,14,1,2).setBackground(WH_D);     // N:O

    var header = sh.getRange(HDR_ROW,11,1,5);
    header.setFontColor('#ffffff').setFontFamily('Roboto').setFontSize(10)
          .setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setFontWeight('bold')
          .setBorder(true,true,true,true,true,true);

    // Данные
    if (rowsNR.length) sh.getRange(DATA_ROW,11,rowsNR.length,5).setValues(rowsNR);
    var totalRows = Math.max(rowsNR.length, 1);
    sh.getRange(DATA_ROW,11,totalRows,5).setFontFamily('Roboto').setFontSize(10)
      .setVerticalAlignment('middle').setHorizontalAlignment('center');

    // K — слева; L — центр; K:L фон #f3f3f3
    sh.getRange(DATA_ROW,11,totalRows,1)
      .setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setBackground('#f3f3f3');
    sh.getRange(DATA_ROW,12,totalRows,1)
      .setHorizontalAlignment('center')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setBackground('#f3f3f3');

    if (rowsNR.length){
      sh.getRange(DATA_ROW,13,rowsNR.length,1).setBackground('#ead1dc'); // M
      sh.getRange(DATA_ROW,14,rowsNR.length,2).setBackground('#fce5cd'); // N:O

      var INT='#,##0;-#,##0;;@';
      sh.getRange(DATA_ROW,12,rowsNR.length,1).setNumberFormat(INT); // L
      sh.getRange(DATA_ROW,14,rowsNR.length,2).setNumberFormat(INT); // N:O
    }

    // Границы
    function rb(c){ sh.getRange(HDR_ROW,c,Math.max(usedRow-HDR_ROW+1,1),1).setBorder(null,null,null,true,null,null,BLACK,SOLID); }
    function lb(c){ sh.getRange(HDR_ROW,c,Math.max(usedRow-HDR_ROW+1,1),1).setBorder(null,true,null,null,null,null,BLACK,SOLID); }
    lb(11); rb(15); // внешний контур K..O
    rb(12); // L | M
    rb(13); // M | N
    if (rowsNR.length) sh.getRange(usedRow,11,1,5).setBorder(null,null,true,null,null,null,BLACK,SOLID);

    // Ширины
    sh.autoResizeColumn(11); sh.setColumnWidth(11, sh.getColumnWidth(11)+50); // K
    sh.setColumnWidths(12,2,110); // L:M
    sh.setColumnWidths(14,2, 60); // N:O
  })();

  // ── Таблица E:I ─────────────────────────────────────────────────────────
  (function draw_EI(){
    var totalData = Math.max(rowsFJ.length, 0);
    var usedRow = totalData ? (DATA_ROW - 1 + rowsFJ.length) : HDR_ROW;

    // Локальный «сброс границ»
    clearBordersLocal_(HDR_ROW, 5, usedRow - HDR_ROW + 1, 5); // E..I

    // Шапка
    sh.getRange(HDR_ROW,5 ).setValue('Товар');      // E2
    sh.getRange(HDR_ROW,6 ).setValue('Бренд');      // F2
    sh.getRange(HDR_ROW,7 ).setValue('Модель');     // G2
    sh.getRange(HDR_ROW,8 ).setValue('Количество'); // H2
    sh.getRange(HDR_ROW,9 ).setValue('Стоимость');  // I2

    sh.getRange(HDR_ROW,5,1,5).setBackground('#434343')
      .setFontColor('#ffffff').setFontWeight('bold')
      .setFontFamily('Roboto').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,true,true);

    // Данные
    if (rowsFJ.length) sh.getRange(DATA_ROW,5,rowsFJ.length,5).setValues(rowsFJ);
    var totalRows = Math.max(rowsFJ.length, 1);
    sh.getRange(DATA_ROW,5,totalRows,5).setFontFamily('Roboto').setFontSize(10)
      .setVerticalAlignment('middle');

    // E — «товар»: серый, жирный, слева; F:G — как артикулы; H:I — числа
    sh.getRange(DATA_ROW,5,totalRows,1)
      .setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setBackground('#cccccc')
      .setFontWeight('bold');

    sh.getRange(DATA_ROW,6,totalRows,2)
      .setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setBackground('#f3f3f3');

    var INT='#,##0;-#,##0;;@';
    sh.getRange(DATA_ROW,8,totalRows,2) // H:I
      .setHorizontalAlignment('center')
      .setBackground('#f3f3f3');
    sh.getRange(DATA_ROW,8,totalRows,1).setNumberFormat(INT);  // H
    sh.getRange(DATA_ROW,9,totalRows,1).setNumberFormat(INT);  // I

    // Контуры
    function rb(c){ sh.getRange(HDR_ROW,c,Math.max(usedRow-HDR_ROW+1,1),1).setBorder(null,null,null,true,null,null,BLACK,SOLID); }
    function lb(c){ sh.getRange(HDR_ROW,c,Math.max(usedRow-HDR_ROW+1,1),1).setBorder(null,true,null,null,null,null,BLACK,SOLID); }
    lb(5); rb(9); // внешний контур E..I
    rb(5);        // E | F
    if (rowsFJ.length) sh.getRange(usedRow,5,1,5).setBorder(null,null,true,null,null,null,BLACK,SOLID);

    // Горизонтальные разделители между брендами
    if (rowsFJData.length > 1){
      for (var i=1;i<rowsFJData.length;i++){
        if (rowsFJData[i].brand !== rowsFJData[i-1].brand){
          var r = DATA_ROW + i;
          sh.getRange(r,5,1,5).setBorder(true,null,null,null,null,null,BLACK,SOLID);
        }
      }
    }

    // Ширины
    sh.autoResizeColumn(5);  sh.setColumnWidth(5,  sh.getColumnWidth(5)+15); // E
    sh.autoResizeColumn(6);  sh.setColumnWidth(6,  sh.getColumnWidth(6)+15); // F
    sh.autoResizeColumn(7);  sh.setColumnWidth(7,  sh.getColumnWidth(7)+15); // G
    sh.setColumnWidth(8,110); // H
    sh.setColumnWidth(9,110); // I
  })();

  // И ещё фикс ширины по заданию (вне таблиц)
  sh.setColumnWidth(18, 55);      // R (повторно, на случай внешних правок)
  sh.setColumnWidths(19,2,65);    // S:T
  sh.setColumnWidths(21,2,65);    // U:V

  // Без закрепления строк
  try { sh.setFrozenRows(0); } catch(_){}
}
