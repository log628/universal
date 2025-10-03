/** =========================
 *  Меню площадки (централизовано)
 * ========================= */

const PARAMS_I2_A1 = 'I2';

function onOpen() {
  try { if (typeof setupCabinetControl_ === 'function') setupCabinetControl_(); } catch (_) {}
  buildPlatformMenu_();
}

function buildPlatformMenu_() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Сменить площадку');

  const ss = SpreadsheetApp.getActive();
  const shParams = ss.getSheetByName(REF.SHEETS.PARAMS);

  const currentRaw = shParams ? String(shParams.getRange(PARAMS_I2_A1).getDisplayValue() || '').trim() : '';
  const norm = normalizePlatform_(currentRaw); // 'OZON' | 'WILDBERRIES' | null

  const label =
    norm === 'OZON'        ? '🟣 Переключить на WILDBERRIES'
  : norm === 'WILDBERRIES' ? '🔵 Переключить на OZON'
                           : '🔁 Переключить: OZON ↔ WILDBERRIES';

  menu.addItem(label, 'menuTogglePlatform_');
  menu.addToUi();
}

function menuTogglePlatform_() {
  const ss = SpreadsheetApp.getActive();

  const shParams = ss.getSheetByName(REF.SHEETS.PARAMS);
  if (!shParams) {
    SpreadsheetApp.getUi().alert(`Не найден лист «${REF.SHEETS.PARAMS}».`);
    return;
  }

  // текущее → следующее значение I2
  const cell = shParams.getRange(PARAMS_I2_A1);
  const cur  = String(cell.getDisplayValue() || '').trim();
  const norm = normalizePlatform_(cur);
  const next = (norm === 'OZON') ? 'WILDBERRIES' : (norm === 'WILDBERRIES' ? 'OZON' : 'OZON');

  cell.setValue(next);
  SpreadsheetApp.flush();

  // перестраиваем дропдаун кабинетов под новую площадку
  try { if (typeof setupCabinetControl_ === 'function') setupCabinetControl_(); } catch (_) {}

  // берём первый кабинет новой площадки из «⚙️ Параметры»
  const firstCab = listCabinetsForPlatform_(next)[0] || '';

  // пишем его в контрол и сразу рендерим
  const shCalc = ss.getSheetByName(REF.SHEETS.CALC);
  if (shCalc && firstCab) {
    const ctrl = shCalc.getRange(REF.CTRL_RANGE_A1);
    ctrl.setValue(firstCab);
    SpreadsheetApp.flush();

    // мгновенный рендер текущей площадки
    try {
      if (typeof runLayoutImmediate === 'function') {
        runLayoutImmediate(firstCab);
      } else if (typeof runLayoutWithDropdownCooldown === 'function') {
        // мягкий фолбэк на старый раннер
        runLayoutWithDropdownCooldown(firstCab);
      } else {
        // последний фолбэк — прямой вызов layout'ов
        if (typeof layoutCalculator === 'function') layoutCalculator(firstCab);
        if (typeof layoutParallel   === 'function') layoutParallel(firstCab);
      }
    } catch (e) {
      ss.toast('Ошибка рендера: ' + ((e && e.message) || e), 'Ошибка', 5);
    }
  } else {
    ss.toast(`Для площадки ${next} не найден ни один кабинет в «${REF.SHEETS.PARAMS}».`, 'Внимание', 5);
  }

  buildPlatformMenu_();
  ss.toast(`Площадка: ${next}${firstCab ? ' — ' + firstCab : ''}`, 'Готово', 3);
}

function listCabinetsForPlatform_(platform /* 'OZON' | 'WILDBERRIES' */) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(REF.SHEETS.PARAMS);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, 4).getDisplayValues(); // A..D
  const platUP = String(platform || '').toUpperCase();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const cab = String(vals[i][0] || '').trim();
    const plt = String(vals[i][3] || '').trim().toUpperCase();
    if (!cab) continue;
    if (platUP && plt === platUP) out.push(cab);
  }
  return Array.from(new Set(out));
}

function normalizePlatform_(raw) {
  const s = String(raw || '').trim();
  if (/^(ozon|oz)$/i.test(s)) return 'OZON';
  if (/^(wildberries|wb)$/i.test(s)) return 'WILDBERRIES';
  return null;
}
