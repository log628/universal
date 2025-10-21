/** ======================================================================
 * setStocks.gs
 * Отправка остатков из листа «🏘️ Собств. склады» в Ozon.
 *
 * Правки: лог всегда со 2-й строки (AD2), без autodetect.
 * ====================================================================== */

/** Точка входа: просто отправить остатки по текущему содержимому листа */
function setStocks() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('🏘️ Собств. склады');
  if (!sh) throw new Error('Лист «🏘️ Собств. склады» не найден');

  // ---------- Утилиты ----------
  var TZ = Session.getScriptTimeZone() || 'Etc/GMT';
  var nowStr = function() { return Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm:ss'); };
  var clean  = function(s){ return String(s == null ? '' : s).trim(); };
  var toNum  = function(v){ var s = String(v == null ? '' : v).replace(/\s+/g,'').replace(',', '.'); var n = Number(s); return isFinite(n) ? n : 0; };
  var firstNonEmpty = function(arr){ for (var i=0;i<arr.length;i++){ var v=clean(arr[i]); if (v) return v; } return ''; };

  // ---------- Лист лога: очистка AD2:AK + заголовки AD1:AK1 ----------
  var logSheet = ss.getSheetByName('🛠 Тех. лог') || ss.insertSheet('🛠 Тех. лог');
  if (logSheet.getMaxColumns() < 37) {
    logSheet.insertColumnsAfter(logSheet.getMaxColumns(), 37 - logSheet.getMaxColumns()); // до AK
  }

  // Заголовки (AD..AK)
  var headers = [['Timestamp','Кабинет','Канал','Warehouse ID','Offer','Qty','Action','Status']];
  logSheet.getRange(1, 30, 1, headers[0].length).setValues(headers).setFontWeight('bold');

  // Очистка содержимого со 2-й строки (строго AD2:AK[низ])
  var rowsToClear = Math.max(0, logSheet.getMaxRows() - 1);
  if (rowsToClear > 0) {
    logSheet.getRange(2, 30, rowsToClear, headers[0].length).clearContent();
  }

// Формат для Warehouse ID (AG) — целое число "0" со 2-й строки вниз
logSheet.getRange(2, 33, Math.max(1, logSheet.getMaxRows() - 1), 1).setNumberFormat('0');








  // Жёсткий указатель строки лога: начинаем всегда с 2
  var LOG_WRITE_ROW = 2;
  function appendLogs(rows){
    if (!rows || !rows.length) return;
    logSheet.getRange(LOG_WRITE_ROW, 30, rows.length, rows[0].length).setValues(rows);
    LOG_WRITE_ROW += rows.length;
  }

  // ---------- 1) Собираем соответствие "кабинет → warehouseId по типам" из B:H ----------
  var maxRows = sh.getMaxRows();
  var bhRange = sh.getRange(1, 2, maxRows, 7); // B:H
  var merges  = bhRange.getMergedRanges();
  var headerRowByCab = new Map(); // FullCabinetName -> headerRow (мердж в 1 строку)

  for (var m=0; m<merges.length; m++){
    var rg = merges[m];
    if (rg.getNumRows() !== 1) continue;
    if (rg.getColumn() < 2 || rg.getColumn() > 8) continue;
    var name = clean(rg.getCell(1,1).getDisplayValue());
    if (name) headerRowByCab.set(name, rg.getRow());
  }

  function getWarehouseMapForCab(cabinetName){
    var map = { Standart:'', Express:'', Comfort:'' }; // тип -> warehouseId
    var row = headerRowByCab.get(cabinetName);
    if (!row) return map;

    // три строки под шапкой
    var types = sh.getRange(row+1, 3, 3, 1).getDisplayValues(); // C
    var ids   = sh.getRange(row+1, 5, 3, 4).getDisplayValues(); // E:H

    for (var i=0;i<3;i++){
      var typ = clean(types[i][0]);
      if (!typ) continue;
      var idRaw = firstNonEmpty(ids[i]); // первый непустой в E..H
      if (!idRaw) continue;
      map[typ] = clean(idRaw);
    }
    return map;
  }

  // ---------- 2) Парсим правый блок N..: определяем кабинеты и индексы их колонок ----------
  var lastCol = sh.getLastColumn();
  var widthN  = Math.max(0, lastCol - 13); // начиная с N (14-я)
  if (widthN <= 0) throw new Error('Справа от N нет данных');

  var row2 = sh.getRange(2, 14, 1, widthN).getDisplayValues()[0]; // N:... (Row2)
  var row3 = sh.getRange(3, 14, 1, widthN).getDisplayValues()[0]; // N:... (Row3)

  var blocks = []; // { baseCol, cabinetName, idxArt, idxS, idxE, idxC }
  for (var c = 14; c <= lastCol; c += 4){
    var i0 = c - 14;
    var cabName = clean(row2[i0] || '');
    var h0 = clean(row3[i0]   || '');
    var h1 = clean(row3[i0+1] || '');
    var h2 = clean(row3[i0+2] || '');
    var h3 = clean(row3[i0+3] || '');
    if (!cabName) continue;
    if (h0 === 'Артикул' && h1 === 'S' && h2 === 'E' && h3 === 'C'){
      blocks.push({ baseCol:c, cabinetName:cabName, idxArt:c, idxS:c+1, idxE:c+2, idxC:c+3 });
    }
  }
  if (!blocks.length) throw new Error('Не найдено ни одного блокa кабинета (N..: Артикул/S/E/C)');

  // ---------- 3) Для каждого кабинета собираем и отправляем S/E/C ----------
  var lastRow = sh.getLastRow();
  var totalSent = 0;

  for (var b=0; b<blocks.length; b++){
    var blk = blocks[b];
    var cabName = blk.cabinetName;

    // Креды: подтянет OZONAPI из «⚙️ Параметры»
    var accounts = OZONAPI.getAccounts();
    if (!accounts[cabName]) {
      // фиксируем, что кабинет не найден в параметрах (строкой с пустыми offer/qty, чтобы не сдвигать логику)
      appendLogs([[ nowStr(), cabName, 'Standart', '', '', 0, 'zero', 'ERR: cabinet not in ⚙️ Параметры' ]]);
      appendLogs([[ nowStr(), cabName, 'Express',  '', '', 0, 'zero', 'ERR: cabinet not in ⚙️ Параметры' ]]);
      appendLogs([[ nowStr(), cabName, 'Comfort',  '', '', 0, 'zero', 'ERR: cabinet not in ⚙️ Параметры' ]]);
      continue;
    }

    // Warehouses по типам
    var wh = getWarehouseMapForCab(cabName); // Standart/Express/Comfort -> id

    // Читаем данные по колонкам
    var height = Math.max(0, lastRow - 3);
    if (height <= 0) continue;

    var arts = sh.getRange(4, blk.idxArt, height, 1).getDisplayValues();
    var sCol = sh.getRange(4, blk.idxS,   height, 1).getDisplayValues();
    var eCol = sh.getRange(4, blk.idxE,   height, 1).getDisplayValues();
    var cCol = sh.getRange(4, blk.idxC,   height, 1).getDisplayValues();

    // Подготовка массивов на отправку по каждому типу
    var payload = { Standart: [], Express: [], Comfort: [] };

    // Проход по строкам
    for (var r=0; r<height; r++){
      var offer = clean(arts[r][0]);
      if (!offer) continue;

      var valS = (sCol[r][0] === '' ? 0 : toNum(sCol[r][0]));
      var valE = (eCol[r][0] === '' ? 0 : toNum(eCol[r][0]));
      var valC = (cCol[r][0] === '' ? 0 : toNum(cCol[r][0]));

      if (wh.Standart) payload.Standart.push({ offer_id: offer, stock: valS, warehouse_id: Number(wh.Standart) });
      if (wh.Express)  payload.Express .push({ offer_id: offer, stock: valE, warehouse_id: Number(wh.Express)  });
      if (wh.Comfort)  payload.Comfort .push({ offer_id: offer, stock: valC, warehouse_id: Number(wh.Comfort)  });
    }

    // Отправка по типам
    function sendType(typeName){
      var warehouseId = wh[typeName];
      if (!warehouseId) return;

      var arr = payload[typeName];
      if (!arr.length) return;

      var api = new OZONAPI(cabName, warehouseId);
      var status = 'OK';
      try {
        api.setStocks(arr); // разбивка по 100 реализована внутри OZONAPI
      } catch (e) {
        status = 'ERR: ' + (e && e.message ? e.message : e);
      }

      // Лог по каждой позиции (Action: set|zero в зависимости от qty), начиная строго с AD2
      var rows = new Array(arr.length);
      for (var i=0;i<arr.length;i++){
        rows[i] = [
          nowStr(),
          cabName,
          typeName,
          String(arr[i].warehouse_id),
          String(arr[i].offer_id),
          Number(arr[i].stock),
          (arr[i].stock > 0 ? 'set' : 'zero'),
          status
        ];
      }
      appendLogs(rows);

      if (status === 'OK') totalSent += arr.length;
    }

    sendType('Standart');
    sendType('Express');
    sendType('Comfort');
  }

  ss.toast('setStocks: отправлено позиций (вкл. нули) — ' + totalSent, 'Ozon', 5);
}

/** Объединённый запуск:
 *  1) runWarehouseFast()
 *  2) ждём 5 сек
 *  3) setStocks()
 */
function setStocksFresh() {
  if (typeof runWarehouseFast === 'function') {
    runWarehouseFast();
  } else {
    // Фолбэк на случай отсутствия функции
    var ss = SpreadsheetApp.getActive();
    ss.toast('Обновление «Доступно»…', 'Склад + СС', 3);
    if (typeof Import_Sklad_GHOnly === 'function') Import_Sklad_GHOnly();
    if (typeof buildOwnWarehouses === 'function') buildOwnWarehouses();
  }

  Utilities.sleep(5000); // 5 секунд
  setStocks();
}
