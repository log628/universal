/** =========================
 *  Меню: Обновить выгрузку  +  [Сменить площадку]
 * ========================= */

const PARAMS_I2_A1 = 'I2';

function onOpen() {
  try { if (typeof setupCabinetControl_ === 'function') setupCabinetControl_(); } catch (_) {}
  buildRefreshMenu_();            // ← новое меню с подпунктами
  buildPlatformMenuBrackets_();   // ← старое меню переключателя, но с названием [Сменить площадку]
}

/** ========== [Сменить площадку] (как у тебя, только другой заголовок) ========== */
function buildPlatformMenuBrackets_() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('[Сменить площадку]');

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

    try {
      if (typeof runLayoutImmediate === 'function') {
        runLayoutImmediate(firstCab);
      } else if (typeof runLayoutWithDropdownCooldown === 'function') {
        runLayoutWithDropdownCooldown(firstCab);
      } else {
        if (typeof layoutCalculator === 'function') layoutCalculator(firstCab);
        if (typeof layoutParallel   === 'function') layoutParallel(firstCab);
      }
    } catch (e) {
      ss.toast('Ошибка рендера: ' + ((e && e.message) || e), 'Ошибка', 5);
    }
  } else {
    ss.toast(`Для площадки ${next} не найден ни один кабинет в «${REF.SHEETS.PARAMS}».`, 'Внимание', 5);
  }

  buildPlatformMenuBrackets_();
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

/** ========== Обновить выгрузку (НОВОЕ меню с точными пунктами) ========== */
function buildRefreshMenu_() {
  const ui = SpreadsheetApp.getUi();
  const m  = ui.createMenu('Обновить выгрузку');

  m
    // 🆔 Артикулы
    .addItem('🆔 Артикулы: Ozon',        'menuRefresh_Arts_OZ_')
    .addItem('🆔 Артикулы: Wildberries', 'menuRefresh_Arts_WB_')
    .addSeparator()
    // 📦 Физ. обороты
    .addItem('📦 Физ. обороты: Ozon',        'menuRefresh_Phys_OZ_')
    .addItem('📦 Физ. обороты: Wildberries', 'menuRefresh_Phys_WB_')
    .addSeparator()
    // 🍔 Склад и Себестоимости
    .addItem('🍔 Склад и Себестоимости', 'menuRefresh_Import_Sklad_')
    .addToUi();
}

/** ==== Хэндлеры пунктов (прямые вызовы твоих функций) ==== */

// 🆔 Артикулы
function menuRefresh_Arts_OZ_() {
  const ss = SpreadsheetApp.getActive();
  if (typeof getREFRESH_OZ === 'function') { getREFRESH_OZ(); ss.toast('Артикулы: Ozon — готово', 'OK', 3); }
  else ss.toast('Функция getREFRESH_OZ не найдена', 'Нет обработчика', 5);
}
function menuRefresh_Arts_WB_() {
  const ss = SpreadsheetApp.getActive();
  if (typeof getREFRESH_WB === 'function') { getREFRESH_WB(); ss.toast('Артикулы: Wildberries — готово', 'OK', 3); }
  else ss.toast('Функция getREFRESH_WB не найдена', 'Нет обработчика', 5);
}

// 📦 Физ. обороты
function menuRefresh_Phys_OZ_() {
  const ss = SpreadsheetApp.getActive();
  if (typeof fiz0_OZ === 'function') { fiz0_OZ(); ss.toast('Физ. обороты: Ozon — готово', 'OK', 3); }
  else ss.toast('Функция fiz0_OZ не найдена', 'Нет обработчика', 5);
}
function menuRefresh_Phys_WB_() {
  const ss = SpreadsheetApp.getActive();
  if (typeof fiz0_WB === 'function') { fiz0_WB(); ss.toast('Физ. обороты: Wildberries — готово', 'OK', 3); }
  else ss.toast('Функция fiz0_WB не найдена', 'Нет обработчика', 5);
}

// 🍔 Склад и Себестоимости
function menuRefresh_Import_Sklad_() {
  const ss = SpreadsheetApp.getActive();
  if (typeof Import_Sklad === 'function') { Import_Sklad(); ss.toast('Склад и Себестоимости — готово', 'OK', 3); }
  else ss.toast('Функция Import_Sklad не найдена', 'Нет обработчика', 5);
}
