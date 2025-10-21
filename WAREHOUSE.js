/**
 * Построение таблицы на листе «🏘️ Собств. склады»
 *
 * Включает:
 * 1) Левый блок K:L — только товары с Доступно>0, заголовки, фон, тонкие границы.
 * 2) Правый блок по кабинетам (N…): заголовки, артикулы, S/E/C по режимам:
 *    - sim01*: всегда S=E=C=10, игнорируя "Доступно" и режим.
 *    - «по наличию»: раздача «Доступно(Товар)» в 3 прохода S→E→C (шаги: quota, ceil(quota/2), ceil(quota/2)).
 *    - «неограниченно»: S=min(Доступно,quota), E=C=ceil(min(Доступно,quota)/2).
 *    Значения 0 → пишем пусто. Непустые S/E/C — заливка #ead1dc.
 *    Все вставленные ячейки — тонкие чёрные границы; периметр каждого кабинета — средней линией;
 *    средняя горизонталь между заголовком (Row3) и данными.
 * 3) Режим «фикс-микс»: по порогам из exports_fm_check → берём строку-ступень (точное совпадение
 *    или ближайший меньший/равный). S=exports_fm_standart[i], E=C=exports_fm_expcomf[i].
 *    Если ни один порог ≤ Доступно — оставляем пусто.
 * 4) Диапазон B:H: под мердж-шапкой с полным именем кабинета (в точном совпадении с N…)
 *    лежат 3 строки: B=флаг, C∈{Standart,Express,Comfort}, D — нужно заполнить:
 *    количеством артикулов этого кабинета, у которых соответствующая колонка (S/E/C) > 0.
 *    E:H — не трогаем (ID склада).
 *
 * ВАЖНО: все ВСТАВЛЯЕМЫЕ ДАННЫЕ — Roboto.
 */
function buildOwnWarehouses() {
  var ss = SpreadsheetApp.getActive();
  var SHEET_OUT = '🏘️ Собств. склады';
  var SHEET_SS  = '🍔 СС';
  var SHEET_PAR = '⚙️ Параметры';
  var SHEET_ART = '[OZ] Артикулы';

  var sh = ss.getSheetByName(SHEET_OUT) || ss.insertSheet(SHEET_OUT);

  // ===== helpers =====
  function flatten2D(arr){var out=[];for(var i=0;i<arr.length;i++)for(var j=0;j<arr[i].length;j++)out.push(arr[i][j]);return out;}
  function cleanStr(s){return String(s==null?'':s).replace(/\s+/g,' ').trim();}
  function toNumber(v){
    if(v==null)return 0; if(typeof v==='number')return isFinite(v)?v:0;
    var s=String(v).trim(); if(!s)return 0;
    s=s.replace(/\u00A0|\u2007|\u202F|\u2009/g,'').replace(/\s+/g,'');
    s=s.replace(/[^0-9.,\-]/g,'');
    if(s.indexOf(',')>-1&&s.indexOf('.')>-1)s=s.replace(/\./g,'').replace(',', '.');
    else if(s.indexOf(',')>-1)s=s.replace(',', '.');
    var n=parseFloat(s); return isFinite(n)?n:0;
  }
  function setThinGrid(r){r.setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.SOLID);}
  function setMediumBox(r){r.setBorder(true,true,true,true,false,false,'#000000',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);}
  function setMediumBottom(r){r.setBorder(null,null,true,null,false,false,'#000000',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);}
  function styleHeader(r,bg,fc){r.setBackground(bg).setFontColor(fc).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Roboto');}
  function clearContents(r){r.clear({contentsOnly:true, formatOnly:false});}
  function clearFills(r){r.setBackground('#ffffff');}
  function clearBorders(r){r.setBorder(false,false,false,false,false,false);} // NEW: сброс границ
  function autoPlus(sh,col,px){try{sh.autoResizeColumn(col);var w=sh.getColumnWidth(col);sh.setColumnWidth(col,Math.max(10,w+(px|0)));}catch(_){}} 
  function articleToProduct(art){var s=cleanStr(art);if(s.length>=3)s=s.substring(3);s=s.replace(/_cat\d+$/i,'');return s;}
  function ceilHalf(x){return Math.ceil(x/2);}

  // ===== режимы / параметры =====
  var regimeRange = ss.getRangeByName('exports_regime');
  var regime = regimeRange ? cleanStr(regimeRange.getDisplayValue()).toLowerCase() : '';
  var quaRange = ss.getRangeByName('exports_qua');
  var quota = quaRange ? Math.max(0, Math.floor(toNumber(quaRange.getValue()))) : 0;

  // Диапазоны фикс-микс
  var fmCheckRange    = ss.getRangeByName('exports_fm_check');
  var fmStandartRange = ss.getRangeByName('exports_fm_standart');
  var fmExpcomfRange  = ss.getRangeByName('exports_fm_expcomf');
  var fmRules = [];
  if (fmCheckRange && fmStandartRange && fmExpcomfRange) {
    var C = flatten2D(fmCheckRange.getValues());
    var S = flatten2D(fmStandartRange.getValues());
    var E = flatten2D(fmExpcomfRange.getValues());
    var n = Math.min(C.length, S.length, E.length);
    for (var i=0;i<n;i++) {
      var chk = toNumber(C[i]);
      var sv  = toNumber(S[i]);
      var ev  = toNumber(E[i]);
      if (isFinite(chk) && chk > 0) fmRules.push({check:chk, s:sv, ec:ev});
    }
  }

  // ===== 1) Левый блок: «Товар/Доступно» (K:L) =====
  var kCol=11,lCol=12;
  var totalRows=sh.getMaxRows(), totalCols=sh.getMaxColumns();
  if (totalRows<2000){sh.insertRowsAfter(totalRows,2000-totalRows); totalRows=sh.getMaxRows();}

  // Заголовки (Roboto)
  sh.getRange(2,kCol,2,1).breakApart().merge().setValue('Товар').setFontFamily('Roboto');
  sh.getRange(2,lCol,2,1).breakApart().merge().setValue('Доступно').setFontFamily('Roboto');
  styleHeader(sh.getRange(2,kCol,2,1),'#b45f06','#ffff00');
  styleHeader(sh.getRange(2,lCol,2,1),'#b45f06','#ffff00');
  setThinGrid(sh.getRange(2,kCol,2,2));

  // Данные из «🍔 СС»: A -> K, G -> L (ТОЛЬКО где Доступно > 0)
  var shSS = ss.getSheetByName(SHEET_SS);
  if(!shSS) throw new Error('Лист «🍔 СС» не найден');
  var lastSS=shSS.getLastRow(), items=[], avail=[];
  if(lastSS>=2){
    var Avals=shSS.getRange(2,1,lastSS-1,1).getDisplayValues(); // Товар (A)
    var Gvals=shSS.getRange(2,7,lastSS-1,1).getDisplayValues(); // Доступно (G)
    for (var r=0;r<Avals.length;r++){
      var t=cleanStr(Avals[r][0]);
      var a=toNumber(Gvals[r][0]);
      if(!t) continue;
      if(a>0){ items.push([t]); avail.push([a]); } // только >0
    }
  }

  // === СОРТИРОВКА ПО «Товар» (K:L) по алфавиту, с русской локалью ===
  if (items.length) {
    var pairs = [];
    for (var i = 0; i < items.length; i++) {
      pairs.push({ t: String(items[i][0] || ''), a: avail[i] ? avail[i][0] : '' });
    }
    pairs.sort(function(p, q) {
      return p.t.localeCompare(q.t, 'ru', { sensitivity: 'base' });
    });
    // раскладываем обратно в items/avail
    items = new Array(pairs.length);
    avail = new Array(pairs.length);
    for (var i = 0; i < pairs.length; i++) {
      items[i] = [pairs[i].t];
      avail[i] = [pairs[i].a];
    }
  }

  // Сброс содержимого и заливок с 4-й строки (левый блок)
  clearContents(sh.getRange(4,kCol,totalRows-3,2));
  clearFills(sh.getRange(4,kCol,totalRows-3,2));

  // Вставка данных в K:L — строго по вставленному диапазону → Roboto
  var rowsKL=items.length;
  if(rowsKL){
    var data2col=[]; for(var i=0;i<rowsKL;i++) data2col.push([items[i][0], (avail[i]?avail[i][0]:'')]);

    var rngKL=sh.getRange(4,kCol,rowsKL,2); // Ровно диапазон вставленных данных

    // 1) значения + оформление
    rngKL.setValues(data2col);
    rngKL.setBackground('#fce5cd');
    setThinGrid(rngKL);
    sh.getRange(4,kCol,rowsKL,1).setHorizontalAlignment('left');
    sh.getRange(4,lCol,rowsKL,1).setHorizontalAlignment('center');

    // 2) ШРИФТ ROBOTO — последним отдельным вызовом, только по вставленному диапазону
    var roboto = SpreadsheetApp.newTextStyle().setFontFamily('Roboto').build();
    rngKL.setTextStyle(roboto);
  }

  // Товар -> Доступно (число; ключ в lower)
  var product2avail=new Map();
  for (var i=0;i<rowsKL;i++){
    var prod=cleanStr(items[i][0]).toLowerCase();
    var avn=toNumber(avail[i][0]);
    product2avail.set(prod,avn);
  }

  // ===== 2) Правая часть: кабинеты =====
  var startCol=14; // N
  var clearCols=Math.max(0,totalCols-(startCol-1));
  if(clearCols>0){
    clearContents(sh.getRange(4,startCol,totalRows-3,clearCols));
    clearFills(sh.getRange(4,startCol,totalRows-3,clearCols));
  }
  if(clearCols>0){ clearContents(sh.getRange(2,startCol,2,clearCols)); }

  // exports_prior → ключи (short без первых 3 символов)
  var priorRange=ss.getRangeByName('exports_prior');
  var priors=priorRange?flatten2D(priorRange.getDisplayValues()).map(cleanStr).filter(Boolean):[];

  // ключ -> полное имя кабинета (только D="OZON")
  var key2full=new Map();
  var shPar=ss.getSheetByName(SHEET_PAR);
  if(shPar){
    var lastPar=shPar.getLastRow();
    if(lastPar>=2){
      var A=shPar.getRange(2,1,lastPar-1,1).getDisplayValues(); // full
      var D=shPar.getRange(2,4,lastPar-1,1).getDisplayValues(); // mp
      var E=shPar.getRange(2,5,lastPar-1,1).getDisplayValues(); // short
      for (var r=0;r<A.length;r++){
        var plat=cleanStr(D[r][0]).toUpperCase(); if(plat!=='OZON') continue;
        var full=cleanStr(A[r][0]); if(!full) continue;
        var shortRaw=String(E[r][0]||'');
        var key=cleanStr(shortRaw.length>=3?shortRaw.substring(3):shortRaw).toLowerCase();
        if(!key) continue;
        if(!key2full.has(key)) key2full.set(key,full);
      }
    }
  }
  var fullCabs=[]; for (var p=0;p<priors.length;p++){ var k=priors[p].toLowerCase(); var f=key2full.get(k); if(f) fullCabs.push(f); }

  // по кабинетам: артикулы (A=полное имя, B=арт)
  var cab2arts=new Map();
  var shArt=ss.getSheetByName(SHEET_ART);
  if(shArt){
    var lastArt=shArt.getLastRow();
    if(lastArt>=2){
      var Aart=shArt.getRange(2,1,lastArt-1,1).getDisplayValues();
      var Bart=shArt.getRange(2,2,lastArt-1,1).getDisplayValues();
      for (var r=0;r<Aart.length;r++){
        var cab=cleanStr(Aart[r][0]), art=cleanStr(Bart[r][0]);
        if(!cab||!art) continue;
        if(!cab2arts.has(cab)) cab2arts.set(cab,[]);
        cab2arts.get(cab).push(art);
      }
    }
  }
  // Сортируем артикулы по каждому кабинету
  for (var ci=0; ci<fullCabs.length; ci++){
    var cab=fullCabs[ci];
    if (cab2arts.has(cab)) cab2arts.get(cab).sort(function(a,b){return a.localeCompare(b,'ru',{sensitivity:'base'});});
  }

  // Заголовки и артикулы в правой части
  var headerBG='#741b47', headerRow2Font='#00ff00', headerRow3Font='#ffff00';
  var cabRows=[], cabArts=[], cabProds=[], cabSEC=[];
  for (var idx=0; idx<fullCabs.length; idx++){
    var fullName=fullCabs[idx];
    var baseCol=startCol+idx*4;

    var rngRow2=sh.getRange(2,baseCol,1,4);
    rngRow2.breakApart().merge().setValue(fullName).setFontFamily('Roboto');
    styleHeader(rngRow2,headerBG,headerRow2Font);

    var rngRow3=sh.getRange(3,baseCol,1,4);
    rngRow3.setValues([['Артикул','S','E','C']]).setFontFamily('Roboto');
    styleHeader(rngRow3,headerBG,headerRow3Font);

    sh.setColumnWidth(baseCol+1,50);
    sh.setColumnWidth(baseCol+2,50);
    sh.setColumnWidth(baseCol+3,50);

    var arts=(cab2arts.get(fullName)||[]);
    cabArts[idx]=arts;
    cabRows[idx]=arts.length;

    // продукты для строк
    var prods=new Array(arts.length);
    for (var irow=0;irow<arts.length;irow++){ prods[irow]=articleToProduct(arts[irow]).toLowerCase(); }
    cabProds[idx]=prods;

    // запишем артикулы (Roboto)
    if (arts.length>0){
      var outArts=new Array(arts.length);
      for (var r=0;r<arts.length;r++) outArts[r]=[arts[r]];
      sh.getRange(4,baseCol,arts.length,1).setValues(outArts).setHorizontalAlignment('left').setFontFamily('Roboto');
      autoPlus(sh, baseCol, 30);
    } else {
      autoPlus(sh, baseCol, 30);
    }

    // инициализация S/E/C пустыми строками
    var sec=new Array(arts.length);
    for (var r=0;r<arts.length;r++) sec[r]=['','',''];
    cabSEC[idx]=sec;
  }

  // ===== NEW: Сброс старых границ в правой части N… на высоту по максимальному числу строк (+ небольшой запас) =====
  var maxCabRows = 0;
  for (var i=0; i<cabRows.length; i++) if (cabRows[i] > maxCabRows) maxCabRows = cabRows[i];
  var clearHeight = 2 + maxCabRows + 5; // Row2-3 (шапки) + данные + небольшой запас
  var rightCols = Math.max(0, sh.getMaxColumns() - (startCol - 1));
  if (rightCols > 0 && clearHeight > 0) {
    clearBorders(sh.getRange(2, startCol, clearHeight, rightCols));
  }

  // ===== 3) Заполнение S/E/C с учётом sim01 и режимов =====
  var allProductsSet=new Set();
  for (var ci=0; ci<fullCabs.length; ci++){
    var prods=cabProds[ci]||[];
    for (var irow=0;irow<prods.length;irow++) if (prods[irow]) allProductsSet.add(prods[irow]);
  }
  var allProducts=Array.from(allProductsSet.values());

  for (var pi=0; pi<allProducts.length; pi++){
    var prod=allProducts[pi];

    // sim01*: всегда S=E=C=10
    if (prod.indexOf('sim01')===0) {
      for (var ci=0; ci<fullCabs.length; ci++){
        var p=cabProds[ci]||[];
        var sec=cabSEC[ci];
        for (var irow=0;irow<p.length;irow++){
          if (p[irow]===prod) { sec[irow][0]=10; sec[irow][1]=10; sec[irow][2]=10; }
        }
      }
      continue;
    }

    var available = product2avail.has(prod) ? toNumber(product2avail.get(prod)) : 0;
    if (available <= 0) continue;

    if (regime === 'по наличию') {
      var remain = available;
      var stepS = quota, stepE = ceilHalf(quota), stepC = ceilHalf(quota);

      // Проход S
      if (remain>0 && stepS>0){
        for (var ci=0; ci<fullCabs.length && remain>0; ci++){
          var pS=cabProds[ci]||[], sS=cabSEC[ci];
          for (var irow=0;irow<pS.length && remain>0;irow++){
            if (pS[irow]===prod){
              var give=Math.min(stepS, remain);
              if (give>0) sS[irow][0]=give;
              remain -= give;
            }
          }
        }
      }
      // Проход E
      if (remain>0 && stepE>0){
        for (var ci=0; ci<fullCabs.length && remain>0; ci++){
          var pE=cabProds[ci]||[], sE=cabSEC[ci];
          for (var irow=0;irow<pE.length && remain>0;irow++){
            if (pE[irow]===prod){
              var give2=Math.min(stepE, remain);
              if (give2>0) sE[irow][1]=give2;
              remain -= give2;
            }
          }
        }
      }
      // Проход C
      if (remain>0 && stepC>0){
        for (var ci=0; ci<fullCabs.length && remain>0; ci++){
          var pC=cabProds[ci]||[], sC=cabSEC[ci];
          for (var irow=0;irow<pC.length && remain>0;irow++){
            if (pC[irow]===prod){
              var give3=Math.min(stepC, remain);
              if (give3>0) sC[irow][2]=give3;
              remain -= give3;
            }
          }
        }
      }

    } else if (regime === 'неограниченно') {
      var capMin = Math.min(available, quota);
      var sVal = (capMin>0) ? capMin : '';
      var eVal = (capMin>0) ? ceilHalf(capMin) : '';
      var cVal = (capMin>0) ? ceilHalf(capMin) : '';
      for (var ci=0; ci<fullCabs.length; ci++){
        var pN=cabProds[ci]||[], sN=cabSEC[ci];
        for (var irow=0;irow<pN.length;irow++){
          if (pN[irow]===prod){ sN[irow][0]=sVal; sN[irow][1]=eVal; sN[irow][2]=cVal; }
        }
      }

    } else if (regime === 'фикс-микс') {
      var bestIdx = -1, bestVal = -Infinity;
      for (var r=0; r<fmRules.length; r++){
        var chk = fmRules[r].check;
        if (chk <= available && chk > bestVal) { bestVal = chk; bestIdx = r; }
      }
      if (bestIdx === -1) continue;

      var sSet = toNumber(fmRules[bestIdx].s);
      var ecSet = toNumber(fmRules[bestIdx].ec);
      var sOut = (sSet>0)?sSet:'';     // 0 → пусто
      var eOut = (ecSet>0)?ecSet:'';   // 0 → пусто
      var cOut = (ecSet>0)?ecSet:'';   // 0 → пусто

      for (var ci=0; ci<fullCabs.length; ci++){
        var pF=cabProds[ci]||[], sF=cabSEC[ci];
        for (var irow=0;irow<pF.length;irow++){
          if (pF[irow]===prod){ sF[irow][0]=sOut; sF[irow][1]=eOut; sF[irow][2]=cOut; }
        }
      }
    }
  }

  // ===== 4) Запись S/E/C, заливки и границы для каждого кабинета =====
  for (var idx=0; idx<fullCabs.length; idx++){
    var baseCol = startCol + idx*4;
    var rows = cabRows[idx];

    if (rows>0) {
      var rngSEC = sh.getRange(4, baseCol+1, rows, 3);
      rngSEC.setValues(cabSEC[idx]).setFontFamily('Roboto');

      // Заливка #ead1dc ТОЛЬКО для непустых значений
      var bg = new Array(rows);
      for (var r=0;r<rows;r++){
        bg[r]=new Array(3);
        for (var c=0;c<3;c++){
          var v = cabSEC[idx][r][c];
          bg[r][c] = (v === '' || v === 0) ? '#ffffff' : '#ead1dc';
        }
      }
      rngSEC.setBackgrounds(bg);

      // Тонкая сетка на весь блок (Row2..Row3+rows)
      var rngBlockAll = sh.getRange(2, baseCol, rows+2, 4);
      setThinGrid(rngBlockAll);
      // Средняя рамка по периметру
      setMediumBox(rngBlockAll);
      // Средняя линия между заголовками и данными (низ Row3)
      setMediumBottom(sh.getRange(3, baseCol, 1, 4));
    } else {
      var rngBlockAll0 = sh.getRange(2, baseCol, 2, 4);
      setThinGrid(rngBlockAll0);
      setMediumBox(rngBlockAll0);
      setMediumBottom(sh.getRange(3, baseCol, 1, 4));
    }
  }

  // ===== 5) Подсчёт и запись D в диапазоне B:H (три строки под шапкой на кабинет) =====
  var cabCounts = new Map(); // fullCabName -> { s:countS, e:countE, c:countC }
  for (var ci=0; ci<fullCabs.length; ci++){
    var sec = cabSEC[ci] || [];
    var cntS=0,cntE=0,cntC=0;
    for (var r=0; r<sec.length; r++){
      if (toNumber(sec[r][0]) > 0) cntS++;
      if (toNumber(sec[r][1]) > 0) cntE++;
      if (toNumber(sec[r][2]) > 0) cntC++;
    }
    cabCounts.set(fullCabs[ci], {s:cntS, e:cntE, c:cntC});
  }

  // Найдём мердж-шапки в B:H и расставим D под ними
  var bhRange2 = sh.getRange(1, 2, sh.getMaxRows(), 7); // B:H
  var merges2 = bhRange2.getMergedRanges();
  var bhHeaders = new Map();
  for (var m=0; m<merges2.length; m++){
    var rg2 = merges2[m];
    if (rg2.getNumRows() !== 1) continue;
    var col2 = rg2.getColumn();
    if (col2 < 2 || col2 > 8) continue;
    var name2 = cleanStr(rg2.getCell(1,1).getDisplayValue());
    if (name2) bhHeaders.set(name2, rg2.getRow());
  }

  bhHeaders.forEach(function(headerRow, cabName){
    if (!cabCounts.has(cabName)) return;
    var types = sh.getRange(headerRow+1, 3, 3, 1).getDisplayValues(); // колонка C
    var outD  = [[0],[0],[0]];
    var counts = cabCounts.get(cabName);
    for (var i=0;i<3;i++){
      var t = cleanStr(types[i][0]).toLowerCase();
      if (t === 'standart') outD[i][0] = counts.s;
      else if (t === 'express') outD[i][0] = counts.e;
      else if (t === 'comfort') outD[i][0] = counts.c;
      else outD[i][0] = 0;
    }
    sh.getRange(headerRow+1, 4, 3, 1).setValues(outD).setFontFamily('Roboto');
  });

  // ===== 6) Подсказка по режимам =====
  if (regime && ['по наличию','неограниченно','фикс-микс'].indexOf(regime)===-1) {
    ss.toast('Поддержаны режимы: «по наличию», «неограниченно», «фикс-микс». Правило sim01 действует всегда (S=E=C=10).', SHEET_OUT, 7);
  }
}
