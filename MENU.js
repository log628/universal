function onOpen() {
  try { if (typeof setupCabinetControl_ === 'function') setupCabinetControl_(); } catch (_) {}
  buildExportMenu_();
  buildImportMenu_();
}

/** ========== 🚀 Экспорт ========== */
function buildExportMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Экспорт')
    .addItem('🔖 Цены [⚖️ Калькулятор]', 'sendPricesFromCalculatorFast')
        .addSeparator()
    .addItem('🥡 Количество [🏘️ Собств. склады]', 'sendPricesFromCalculatorFast')      
            .addSeparator()
    .addItem('📬 Рецепт [🎏 Форкаст]', 'emailForecastAsXlsx')  
    .addToUi();
}

/** ========== 🛸 Импорт (с блоком 🏷️ Цены) ========== */
function buildImportMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('🛸 Импорт')
    // 🆔 Артикулы
    .addItem('🆔 Артикулы: Ozon',        'getREFRESH_OZ')
    .addItem('🆔 Артикулы: Wildberries', 'getREFRESH_WB')
    .addSeparator()
    // 🏷️ Цены
    .addItem('🏷️ Цены: Ozon',        'getREFRESHprices_OZ')
    .addItem('🏷️ Цены: Wildberries', 'getREFRESHprices_WB')
    .addSeparator()
    // 📦 Физ. обороты — ИМЕНА ИСПРАВЛЕНЫ
    .addItem('📦 Физ. обороты: Ozon',        'getFiz_OZ')
    .addItem('📦 Физ. обороты: Wildberries', 'getFiz_WB')
    .addSeparator()
    // 🍔 Склад и Себестоимости
    .addItem('🍔 Склад и Себестоимости', 'Import_Sklad')
    .addToUi();
}
