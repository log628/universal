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

  // ────────────────────────────────────────────────────────────────────────
  // Ф И Л Ь Т Р Ы  (вкладка 🎏 Форкаст, столбцы B:C)
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

  // ────────────────────────────────────────────────────────────────────────
  // П А Р А М Е Т Р Ы
  // ────────────────────────────────────────────────────────────────────────
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
  // Ч Т Е Н И Е  «🍔 СС»: ТОВАРЫ, ФЛАГ «НЕ ЗАКУПАЕТСЯ», БРЕНД/МОДЕЛЬ/ВАЛЮТА,
  // КУРСЫ, КОМПЛЕКТЫ O:P:Q
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

    var vals = s.getRange(2,1,lr-1,13).getDisplayValues(); // A:M
    for (var i=0;i<vals.length;i++){
      var tv = String(vals[i][0]||'').trim(); if (!tv) continue;      // A
      var brand = String(vals[i][1]||'').trim();                      // B
      var model = String(vals[i][2]||'').trim();                      // C
      var ccCur = num(vals[i][3]);                                    // D
      var currency = String(vals[i][4] || '').trim();                 // E
      var nal=num(vals[i][6]), vput=num(vals[i][7]), vpo=num(vals[i][8]), vpw=num(vals[i][9]); // G,H,I,J
      var notBuy = norm(vals[i][9]) === 'да'; // J = "Не закупается"
      if (notBuy) out.notBuySet.add(tv);
      out.goods.set(tv, {
        brand: brand, model: model, ccCur: isFinite(ccCur)?ccCur:0, currency: currency,
        nal: isFinite(nal)?nal:0, vput: isFinite(vput)?vput:0, vpostSum: (isFinite(vpo)?vpo:0) + (isFinite(vpw)?vpw:0),
        notBuy: notBuy
      });
    }

    // Курсы L:M
    var valsLM = s.getRange(2,12,lr-1,2).getDisplayValues(); // L,M
    for (var j=0;j<valsLM.length;j++){
      var cname = (valsLM[j][0]||'').trim();
      var rate  = num(valsLM[j][1]);
      if (!cname) continue;
      out.rates.set(norm(cname), isFinite(rate)?rate:0);
    }

    // Комплекты O:P:Q
    var lastCol = s.getLastColumn();
    if (lastCol >= 17){
      var valsOPQ = s.getRange(2,15,lr-1,3).getDisplayValues(); // O:P:Q
      for (var k=0;k<valsOPQ.length;k++){
        var kit = String(valsOPQ[k][0]||'').trim();
        var comp= String(valsOPQ[k][1]||'').trim();
        var coef= num(valsOPQ[k][2]);
        if (!kit || !comp) continue;
        var c = isFinite(coef) ? coef : 0;
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
  // Ч Т Е Н И Е  А Р Т И К У Л О В  + К А Т Е Г О Р И И (СВ. КАТ.)
  // ────────────────────────────────────────────────────────────────────────
  function readArticlesWithCategory_(sheetName){
    var s = ss.getSheetByName(sheetName);
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
  // Г Р У П П И Р О В К А  П О  Т О В А Р А М  (+ фильтр "не закупается")
  // ────────────────────────────────────────────────────────────────────────
  var byTovar = new Map();
  var tovarCats = new Map();
  function add(platformTag, rec){
    var tv = tovarFromArticle(platformTag, rec.art);
    if (!tv) return;
    // Фильтр 1: «Не закупается» — исключаем полностью из ВСЕХ таблиц
    if (SS.notBuySet.has(tv)) return;

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

  // Фильтр 4: CHINAONLY (валюта не-рубль)
  var TOVARS_CHINA = TOVARS_ALL.filter(function(tv){
    if (!CHINAONLY) return true;
    var rec = SS.goods.get(tv);
    if (!rec) return false;
    if (!rec.currency) return false;
    return !isRubleCurrency_(rec.currency);
  });

  // ────────────────────────────────────────────────────────────────────────
  // Ф И З . О Б О Р О Т  (по артикулу, раздельно OZ/WB)
  // ────────────────────────────────────────────────────────────────────────
  function readFizMapsByArticle_(sheetName, enabled){
    var maps = { fbo:new Map(), spd:new Map() };
    if (!enabled) return maps;
    var s = ss.getSheetByName(sheetName);
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
  // Р А С Ч Ё Т  «П О Т Р Е Б Н О С Т И»  Д Л Я  А Р Т И К У Л О В  (T:Z)
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
  // П О Д Г О Т О В К А  Д А Н Н Ы Х  T:Z  (товары+артикулы с агрегированием)
  // ────────────────────────────────────────────────────────────────────────
  var rowsTZ = [];
  var productIdx = [];
  var sumNeedByTovar_raw = new Map(); // сумма Z по артикульным строкам товара (до комплектов)

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
      rowsTZ.push(['  '+tv, (ozCount+' | '+wbCount), 0,0,0,0,0]); // T..Z
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
  // К О М П Л Е К Т Ы  и  «В Т О Р О Й  Т И П»
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

  // selfCoef для «второго типа»
  var selfCoef = new Map(); // tv -> sum Q для строк O==P==tv
  (SS.kits||[]).forEach(function(edge){
    if (edge.kit && edge.comp && edge.kit === edge.comp){
      selfCoef.set(edge.kit, (selfCoef.get(edge.kit)||0) + edge.coef);
    }
  });

  // Добавки от комплектов к компонентам (исключая self, он идёт в base)
  var addFromKits = new Map();
  (SS.kits||[]).forEach(function(edge){
    var kit = edge.kit, comp = edge.comp, c = edge.coef;
    if (!kit || !comp || c<=0) return;
    if (kit === comp) return; // self — в base для второго типа

    if (!sumNeedByTovar_raw.has(kit)) return; // комплект не прошёл фильтры → нет вклада
    var zKit = sumNeedByTovar_raw.get(kit) || 0;
    if (zKit <= 0) return;

    addFromKits.set(comp, (addFromKits.get(comp)||0) + zKit * c);
  });

  // ────────────────────────────────────────────────────────────────────────
  // Т А Б Л И Ц А  N:R — по товарам (без «обычных комплектов»)
  // ────────────────────────────────────────────────────────────────────────
  var rowsNR = [];
  if (!nothingToCollect){
    var listForNR = T_sorted.filter(function(tv){
      return !isBrandKit(tv); // обычные комплекты исключаем
    });

    for (var i=0;i<listForNR.length;i++){
      var tv = listForNR[i];
      var ssrec = (SS.goods.get(tv) || { nal:0, vput:0 });

      var zRaw = sumNeedByTovar_raw.get(tv) || 0; // базовая из T:Z
      var base;
      if (compSecondType.has(tv)){
        var sc = selfCoef.get(tv) || 0;
        base = zRaw * sc; // база для второго типа — Z * coef_self
      } else {
        base = zRaw;
      }

      var plusFromKits = addFromKits.get(tv) || 0;
      var P_total = base + plusFromKits;

      // Строковое P — БЕЗ ПРОБЕЛА перед "+"
      var P_disp = '';
      if (base > 0 && plusFromKits > 0) P_disp = String(base) + '+' + String(plusFromKits);
      else if (base <= 0 && plusFromKits > 0) P_disp = '+' + String(plusFromKits);
      else if (base > 0 && plusFromKits <= 0) P_disp = String(base);
      else P_disp = ''; // обе 0 → пусто

      // К закупу:
      var baseKup = Math.max(0, (P_total + MINIMAL) - (ssrec.nal + ssrec.vput));
      var kup  = (baseKup < 3) ? 0 : ceilToStep(baseKup, ROUNDSTEP);

      rowsNR.push(['  '+tv, kup, P_disp, ssrec.nal||0, ssrec.vput||0]);
    }
  }

  // ────────────────────────────────────────────────────────────────────────
  // Т А Б Л И Ц А  F:J — к закупу (из N:R, где O>0)
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
  // Х Р О Н О Л О Г И Я  П Е Р Е Р И С О В К И
  // ────────────────────────────────────────────────────────────────────────
  var lastRowSheet = Math.max(sh.getMaxRows(), 2);
  var WHITE = '#ffffff';
  var SOLID = SpreadsheetApp.BorderStyle.SOLID;

  // (1) Очистка контента
  sh.getRange(1,14,lastRowSheet,5).clearContent(); // N:R
  sh.getRange(1,20,lastRowSheet,7).clearContent(); // T:Z
  sh.getRange(1,6 ,lastRowSheet,5).clearContent(); // F:J

  // (2) Обеление и белые границы
  sh.getRange(1,14,lastRowSheet,5).setBackground(WHITE).setBorder(true,true,true,true,true,true,WHITE,SOLID);
  sh.getRange(1,20,lastRowSheet,7).setBackground(WHITE).setBorder(true,true,true,true,true,true,WHITE,SOLID);
  sh.getRange(1,6 ,lastRowSheet,5).setBackground(WHITE).setBorder(true,true,true,true,true,true,WHITE,SOLID);

  // (3) Сброс границ под будущие области
  var tzRows = Math.max(rowsTZ.length, 1), nrRows = Math.max(rowsNR.length, 1), fjRows = Math.max(rowsFJ.length, 1);
  var tzLast = 1 + tzRows; // header 1 + data
  var nrLast = 1 + nrRows;
  var fjLast = 1 + fjRows;

  sh.getRange(1,20,Math.max(tzLast,1),7).setBorder(false,false,false,false,false,false,null,null); // T:Z
  sh.getRange(1,14,Math.max(nrLast,1),5).setBorder(false,false,false,false,false,false,null,null); // N:R
  sh.getRange(1,6 ,Math.max(fjLast,1),5).setBorder(false,false,false,false,false,false,null,null); // F:J

  // (4) Рисуем таблицы

  // ── Таблица T:Z ─────────────────────────────────────────────────────────
  (function draw_TZ(){
    var DATA_START=2, totalRows = Math.max(rowsTZ.length, 1), usedRow = DATA_START - 1 + rowsTZ.length;

    // Шапка
    sh.getRange(1,20).setValue('Товар / Артикул');   // T
    sh.getRange(1,21).setValue('МП');                 // U
    sh.getRange(1,22).setValue('Остаток FBO');        // V:W merged
    sh.getRange(1,24).setValue('Скорость');           // X:Y merged
    sh.getRange(1,26).setValue('Потребность');        // Z
    sh.getRange(1,22,1,2).merge(); // V1:W1
    sh.getRange(1,24,1,2).merge(); // X1:Y1

    var WHITE='#ffffff', DARK='#434343', FBO_D='#274e13', SPD_D='#1c4587', NEED_D='#741b47';
    sh.getRange(1,20,1,2).setBackground(DARK);
    sh.getRange(1,22,1,2).setBackground(FBO_D);
    sh.getRange(1,24,1,2).setBackground(SPD_D);
    sh.getRange(1,26,1,1).setBackground(NEED_D);

    var header = sh.getRange(1,20,1,7);
    header.setFontColor(WHITE).setFontFamily('Roboto').setFontSize(12)
          .setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setFontWeight('bold')
          .setBorder(true,true,true,true,true,true);

    if (rowsTZ.length) sh.getRange(DATA_START,20,rowsTZ.length,7).setValues(rowsTZ);
    sh.getRange(DATA_START,20,totalRows,7).setFontFamily('Roboto').setFontSize(12)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.getRange(DATA_START,20,totalRows,1)
      .setHorizontalAlignment('left').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    if (totalRows>0){
      sh.getRange(DATA_START,22,totalRows,2).setBackground('#d9ead3'); // V:W
      sh.getRange(DATA_START,24,totalRows,2).setBackground('#c9daf8'); // X:Y
      sh.getRange(DATA_START,26,totalRows,1).setBackground('#ead1dc'); // Z
    }

    var INT='#,##0;-#,##0;;@', DEC2='#,##0.00;-#,##0.00;;@';
    if (totalRows>0){
      sh.getRange(DATA_START,22,totalRows,2).setNumberFormat(INT);
      sh.getRange(DATA_START,24,totalRows,2).setNumberFormat(DEC2);
      sh.getRange(DATA_START,26,totalRows,1).setNumberFormat(INT);
    }

    // Сброс стиля Z перед точечной раскраской (важно для повторных запусков)
    if (totalRows>0){
      sh.getRange(DATA_START,26,totalRows,1).setFontColor('#000000').setFontWeight('normal');
    }

    var BLACK='#000000', SOLID=SpreadsheetApp.BorderStyle.SOLID;
    var prodExcelRows = new Set(productIdx.map(function(i){ return DATA_START + i; }));

    for (var i=0;i<rowsTZ.length;i++){
      var r = DATA_START + i;
      if (prodExcelRows.has(r)){
        // строка ТОВАРА: общий серый фон, жирный, чёрный текст (в т.ч. Z)
        sh.getRange(r,20,1,7).setBackground('#cccccc').setFontWeight('bold')
          .setBorder(true,null,true,null,null,null,BLACK,SOLID);
        sh.getRange(r,26,1,1).setFontColor('#000000'); // Z чёрный
      } else {
        // строка АРТИКУЛА
        sh.getRange(r,20,1,2).setBackground('#f3f3f3'); // T:U
        var row = rowsTZ[i];
        var FBO_OZ=Number(row[2]), FBO_WB=Number(row[3]), SPD_OZ=Number(row[4]), SPD_WB=Number(row[5]), NEED=Number(row[6]);
        if (FBO_OZ===0 && FBO_WB===0 && SPD_OZ===0 && SPD_WB===0) {
          sh.getRange(r,26).setFontWeight('bold').setFontColor('#e69138');
        } else if (NEED>0) {
          sh.getRange(r,26).setFontWeight('bold').setFontColor('#0000ff');
        }
      }
    }

    function rb(c){ sh.getRange(1,c,Math.max(usedRow,1),1).setBorder(null,null,null,true,null,null,BLACK,SOLID); }
    function lb(c){ sh.getRange(1,c,Math.max(usedRow,1),1).setBorder(null,true,null,null,null,null,BLACK,SOLID); }
    lb(20); rb(26); // внешний
    rb(21);         // U | V
    rb(23);         // W | X
    rb(25);         // Y | Z
    if (rowsTZ.length) sh.getRange(usedRow,20,1,7).setBorder(null,null,true,null,null,null,BLACK,SOLID);

    sh.autoResizeColumn(20); sh.setColumnWidth(20, sh.getColumnWidth(20)+50); // T
    sh.setColumnWidth(21, 90);   // U
    sh.setColumnWidths(22,2,75); // V:W
    sh.setColumnWidths(24,2,75); // X:Y
    sh.setColumnWidth(26,110);   // Z
  })();

  // ── Таблица N:R ─────────────────────────────────────────────────────────
  (function draw_NR(){
    var DATA_START=2, totalRows = Math.max(rowsNR.length, 1), usedRow = DATA_START - 1 + rowsNR.length;

    sh.getRange(1,14).setValue('Товар');       // N
    sh.getRange(1,15).setValue('К закупу');    // O
    sh.getRange(1,16).setValue('Потребность'); // P (строковое "base+add")
    sh.getRange(1,17).setValue('Налич');       // Q
    sh.getRange(1,18).setValue('Путь');        // R

    var WHITE='#ffffff', DARK='#434343', NEED_D='#741b47', WH_D='#783f04';
    sh.getRange(1,14,1,2).setBackground(DARK);     // N:O
    sh.getRange(1,16,1,1).setBackground(NEED_D);   // P
    sh.getRange(1,17,1,2).setBackground(WH_D);     // Q:R

    var header = sh.getRange(1,14,1,5);
    header.setFontColor(WHITE).setFontFamily('Roboto').setFontSize(12)
          .setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setFontWeight('bold')
          .setBorder(true,true,true,true,true,true);

    if (rowsNR.length) sh.getRange(DATA_START,14,rowsNR.length,5).setValues(rowsNR);
    sh.getRange(DATA_START,14,totalRows,5).setFontFamily('Roboto').setFontSize(12)
      .setVerticalAlignment('middle').setHorizontalAlignment('center');

    // N — слева, обрезать; O — центр; N:O фон как для артикулов (#f3f3f3)
    sh.getRange(DATA_START,14,totalRows,1)
      .setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setBackground('#f3f3f3');
    sh.getRange(DATA_START,15,totalRows,1)
      .setHorizontalAlignment('center')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setBackground('#f3f3f3');

    if (rowsNR.length){
      sh.getRange(DATA_START,16,rowsNR.length,1).setBackground('#ead1dc'); // P
      sh.getRange(DATA_START,17,rowsNR.length,2).setBackground('#fce5cd'); // Q:R

      var INT='#,##0;-#,##0;;@';
      sh.getRange(DATA_START,15,rowsNR.length,1).setNumberFormat(INT); // O
      sh.getRange(DATA_START,17,rowsNR.length,2).setNumberFormat(INT); // Q:R
    }

    // Границы: внешний + O|P и P|Q
    var BLACK='#000000', SOLID=SpreadsheetApp.BorderStyle.SOLID;
    function rb(c){ sh.getRange(1,c,Math.max(usedRow,1),1).setBorder(null,null,null,true,null,null,BLACK,SOLID); }
    function lb(c){ sh.getRange(1,c,Math.max(usedRow,1),1).setBorder(null,true,null,null,null,null,BLACK,SOLID); }
    lb(14); rb(18);
    rb(15); // O | P
    rb(16); // P | Q
    if (rowsNR.length) sh.getRange(usedRow,14,1,5).setBorder(null,null,true,null,null,null,BLACK,SOLID);

    sh.autoResizeColumn(14); sh.setColumnWidth(14, sh.getColumnWidth(14)+50); // N
    sh.setColumnWidths(15,2,110); // O:P
    sh.setColumnWidths(17,2,60);  // Q:R
  })();

  // ── Таблица F:J ─────────────────────────────────────────────────────────
  (function draw_FJ(){
    var DATA_START=2, totalRows = Math.max(rowsFJ.length, 1), usedRow = DATA_START - 1 + rowsFJ.length;

    sh.getRange(1,6 ).setValue('Товар');     // F
    sh.getRange(1,7 ).setValue('Бренд');     // G
    sh.getRange(1,8 ).setValue('Модель');    // H
    sh.getRange(1,9 ).setValue('Количество');// I
    sh.getRange(1,10).setValue('Стоимость'); // J

    var WHITE='#ffffff', DARK='#434343';
    sh.getRange(1,6,1,5).setBackground(DARK)
      .setFontColor(WHITE).setFontWeight('bold')
      .setFontFamily('Roboto').setFontSize(12)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,true,true);

    if (rowsFJ.length) sh.getRange(DATA_START,6,rowsFJ.length,5).setValues(rowsFJ);
    sh.getRange(DATA_START,6,totalRows,5).setFontFamily('Roboto').setFontSize(12)
      .setVerticalAlignment('middle');

    // F — «товар»: серый, жирный, слева, обрезать
    sh.getRange(DATA_START,6,totalRows,1)
      .setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setBackground('#cccccc')
      .setFontWeight('bold');

    // G:H — «как артикулы»: слева, обрезать, фон #f3f3f3
    sh.getRange(DATA_START,7,totalRows,2)
      .setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setBackground('#f3f3f3');

    // I:J — центр, фон как G:H, целые числа
    var INT='#,##0;-#,##0;;@';
    sh.getRange(DATA_START,9,totalRows,2)
      .setHorizontalAlignment('center')
      .setBackground('#f3f3f3');
    sh.getRange(DATA_START,9,totalRows,1).setNumberFormat(INT);  // I
    sh.getRange(DATA_START,10,totalRows,1).setNumberFormat(INT); // J

    // Внешний контур + вертикальная грань F|G
    var BLACK='#000000', SOLID=SpreadsheetApp.BorderStyle.SOLID;
    function rb(c){ sh.getRange(1,c,Math.max(usedRow,1),1).setBorder(null,null,null,true,null,null,BLACK,SOLID); }
    function lb(c){ sh.getRange(1,c,Math.max(usedRow,1),1).setBorder(null,true,null,null,null,null,BLACK,SOLID); }
    lb(6); rb(10);
    rb(6); // F | G
    if (rowsFJ.length) sh.getRange(usedRow,6,1,5).setBorder(null,null,true,null,null,null,BLACK,SOLID);

    // Горизонтальные разделители между брендами
    if (rowsFJData.length > 1){
      for (var i=1;i<rowsFJData.length;i++){
        if (rowsFJData[i].brand !== rowsFJData[i-1].brand){
          var r = DATA_START + i;
          sh.getRange(r,6,1,5).setBorder(true,null,null,null,null,null,BLACK,SOLID);
        }
      }
    }

    // Ширины: F:G:H авто +15; I:J по 110
    sh.autoResizeColumn(6);  sh.setColumnWidth(6,  sh.getColumnWidth(6)+15);
    sh.autoResizeColumn(7);  sh.setColumnWidth(7,  sh.getColumnWidth(7)+15);
    sh.autoResizeColumn(8);  sh.setColumnWidth(8,  sh.getColumnWidth(8)+15);
    sh.setColumnWidth(9, 110);
    sh.setColumnWidth(10,110);
  })();

  // Без закрепления строк
  try { sh.setFrozenRows(0); } catch(_){}
}
