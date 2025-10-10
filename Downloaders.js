/**
 * Экспорт таблицы F:J (🎏 Форкаст) в XLSX и отправка на почту.
 * Ничего не сохраняет на Диск: создаётся временная таблица → экспортируется → удаляется.
 *
 * @param {string=} recipient Email получателя (опционально).
 *                            Если не задан: берём Session.getActiveUser().getEmail(),
 *                            если пусто — email владельца текущей таблицы.
 */
function emailForecastFJ_XLSX(recipient) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF && REF.SHEETS ? REF.SHEETS.FORECAST : '🎏 Форкаст');
  if (!sh) throw new Error('Лист "🎏 Форкаст" не найден');

  // Получатель
  var to = String(recipient || '').trim();
  if (!to) {
    to = Session.getActiveUser().getEmail() || '';
  }
  if (!to) {
    try {
      to = DriveApp.getFileById(ss.getId()).getOwner().getEmail();
    } catch (_) {}
  }
  if (!to) throw new Error('Не удалось определить email получателя. Передайте email параметром.');

  // Читаем F:J (заголовок + данные)
  var START_COL = 6; // F
  var NUM_COLS  = 5; // F..J
  var lastRow   = sh.getLastRow();
  if (lastRow < 2) throw new Error('Нет данных для экспорта');

  var all = sh.getRange(1, START_COL, lastRow, NUM_COLS).getDisplayValues();
  function isRowEmpty(rowArr){ for (var c=0;c<NUM_COLS;c++){ if (String(rowArr[c]||'').trim()!=='') return false; } return true; }

  var lastNonEmpty = 1; // хотя бы заголовок
  for (var r = all.length - 1; r >= 1; r--) {
    if (!isRowEmpty(all[r])) { lastNonEmpty = r; break; }
  }
  if (lastNonEmpty < 1) throw new Error('Нет данных F:J для экспорта');
  var data = all.slice(0, lastNonEmpty + 1);

  // Временная таблица
  var tz = ss.getSpreadsheetTimeZone() || 'Etc/GMT';
  var stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  var exportName = 'Forecast F-J ' + stamp;

  var tempSS = SpreadsheetApp.create(exportName);
  var dst    = tempSS.getSheets()[0];
  dst.setName('Export');
  dst.getRange(1,1,data.length,data[0].length).setValues(data);

  // Мини-оформление шапки и форматов (не обязательно для письма, но приятно)
  var DARK='#434343', WHITE='#ffffff';
  dst.getRange(1,1,1,NUM_COLS)
     .setBackground(DARK).setFontColor(WHITE)
     .setFontFamily('Roboto').setFontSize(12).setFontWeight('bold')
     .setHorizontalAlignment('center').setVerticalAlignment('middle');
  if (data.length > 1) {
    var rows = data.length - 1;
    dst.getRange(2,1,rows,1).setHorizontalAlignment('left').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    dst.getRange(2,2,rows,2).setHorizontalAlignment('left').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    dst.getRange(2,4,rows,2).setHorizontalAlignment('center');
    var INT = '#,##0;-#,##0;;@';
    dst.getRange(2,4,rows,1).setNumberFormat(INT);
    dst.getRange(2,5,rows,1).setNumberFormat(INT);
  }

  // Экспорт в XLSX (в память), без сохранения файла на Диск
  var mimeXlsx = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var blob;
  try {
    var resp = Drive.Files.export(tempSS.getId(), mimeXlsx, { alt: 'media' }); // Advanced Drive Service
    blob = (resp && typeof resp.getBlob === 'function') ? resp.getBlob() : resp;
    if (!blob) throw new Error('Empty blob from Drive.Files.export');
  } catch (e) {
    // Fallback через Drive v3 (UrlFetch)
    var token = ScriptApp.getOAuthToken();
    var fetchUrl = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(tempSS.getId()) +
                   '/export?mimeType=' + encodeURIComponent(mimeXlsx);
    var fetchResp = UrlFetchApp.fetch(fetchUrl, { headers: { Authorization: 'Bearer ' + token } });
    blob = fetchResp.getBlob();
  }
  blob.setName(exportName + '.xlsx');

  // Письмо с вложением
  var subj = exportName;
  var body = 'Отчёт F:J во вложении.\n\nСгенерировано автоматически ' + stamp;
  MailApp.sendEmail({
    to: to,
    subject: subj,
    body: body,
    attachments: [blob]
  });

  // Удалить временную таблицу
  try { Drive.Files.remove(tempSS.getId()); } catch(_){}

  // Для удобства — в лог
  Logger.log('Отправлено на: ' + to + ' | файл: ' + blob.getName());
}
