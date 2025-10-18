function emailForecastAsXlsx() {
  const SRC_SHEET_NAME = '🎏 Форкаст';
  const START_COL = 5;   // E
  const END_COL   = 9;   // I
  const START_ROW = 2;   // заголовки на 2-й строке

  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SRC_SHEET_NAME);
  if (!src) throw new Error('Лист "🎏 Форкаст" не найден');

  const lastRow = getLastDataRowInBlock_(src, START_ROW, START_COL, END_COL - START_COL + 1);
  if (lastRow < START_ROW) {
    ss.toast('В диапазоне E2:I нет данных для экспорта', 'Экспорт Forecast', 5);
    return;
  }

  // 1) Временная таблица и копия листа
  const temp = SpreadsheetApp.create('TEMP_Export_Forecast_E2-I');
  const copied = src.copyTo(temp).setName('Export');
  const def = temp.getSheets()[0];
  if (def.getSheetId() !== copied.getSheetId()) temp.deleteSheet(def);

  // 2) Обрезка по столбцам: оставить E:I (станут A:E)
  const maxCols = copied.getMaxColumns();
  if (maxCols > END_COL) copied.deleteColumns(END_COL + 1, maxCols - END_COL);
  if (START_COL > 1)     copied.deleteColumns(1, START_COL - 1);

  // 3) Обрезка по строкам: удалить ниже последней строки данных
  const maxRows = copied.getMaxRows();
  if (maxRows > lastRow) copied.deleteRows(lastRow + 1, maxRows - lastRow);

  // 4) Удаляем ПЕРВУЮ строку ⇒ заголовки становятся на 1-й строке
  copied.deleteRow(1);

  // 5) Зафиксировать только значения (сохранив оформление)
  const used = copied.getDataRange();
  used.copyTo(used, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  // 6) Фризим 1 строку (шапку)
  copied.setFrozenRows(1);

  SpreadsheetApp.flush();

  // 7) Экспорт .xlsx через UrlFetchApp
  const tz = Session.getScriptTimeZone() || 'Asia/Almaty';
  const now_ddMM = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm');
  const now_file = Utilities.formatDate(new Date(), tz, 'ddMMyy-HHmm');

  const filename = `Forecast_${now_file}.xlsx`;
  const subject  = `🎏 Форкаст — К закупу (${now_ddMM})`;
  const blob     = exportSheetToXlsxBlob_(temp.getId(), filename);

  // 8) Отправка письма запускающему пользователю
  const to = detectRunnerEmail_();
  if (!to) throw new Error('Не удалось определить email запускающего пользователя.');

  MailApp.sendEmail({
    to,
    subject,
    body: '',                // пустое тело письма
    attachments: [blob],
  });

  // 9) Удаляем временный файл
  DriveApp.getFileById(temp.getId()).setTrashed(true);
}

/** Экспортирует Google Sheet в XLSX через OAuth. */
function exportSheetToXlsxBlob_(spreadsheetId, filename) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error(`Export failed: HTTP ${res.getResponseCode()} — ${res.getContentText().slice(0, 300)}`);
  }
  const blob = res.getBlob();
  blob.setName(filename);
  return blob;
}

/** Последний непустой ряд внутри прямоугольного блока (по ЛЮБОМУ столбцу блока). */
function getLastDataRowInBlock_(sheet, startRow, startCol, numCols) {
  const sheetLast = sheet.getLastRow();
  if (sheetLast < startRow) return startRow - 1;
  const numRows = sheetLast - startRow + 1;
  const values  = sheet.getRange(startRow, startCol, numRows, numCols).getDisplayValues();
  let tail = values.length;
  while (tail > 0) {
    const row = values[tail - 1];
    if (row.some(v => String(v).trim() !== '')) break;
    tail--;
  }
  return (tail === 0) ? (startRow - 1) : (startRow + tail - 1);
}

/** Почта запускающего: ActiveUser → EffectiveUser → владелец файла. */
function detectRunnerEmail_() {
  try {
    const a = Session.getActiveUser().getEmail && Session.getActiveUser().getEmail();
    if (a) return a;
  } catch (_) {}
  try {
    const e = Session.getEffectiveUser().getEmail && Session.getEffectiveUser().getEmail();
    if (e) return e;
  } catch (_) {}
  try {
    const owner = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getOwner();
    if (owner && owner.getEmail) return owner.getEmail();
  } catch (_) {}
  return '';
}
