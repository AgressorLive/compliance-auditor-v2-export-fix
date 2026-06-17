/**
 * ================================================================
 *  ADS APPEAL MAIL IMPORTER — Standalone Google Apps Script
 * ================================================================
 *
 *  Призначення:
 *    Моніторить пошту щогодини, шукає листи від Google Ads:
 *
 *    Метод 1 — Маркерні картинки:
 *      sucess.png  → Апеляція ВІДНОВЛЕНА
 *      cancel.png  → Апеляція ВІДХИЛЕНА
 *
 *    Метод 2 — Текстовий аналіз з перекладом (LanguageApp):
 *      "reinstated your account"        → ВІДНОВЛЕНА
 *      "account has been suspended"     → ВІДХИЛЕНА
 *      (працює на будь-якій мові листа)
 *
 *    Результати пишуться в аркуш "Mail Junk" спільної таблиці.
 *
 *  Розгортання:
 *    1. script.google.com → New project → встав цей файл
 *    2. Заповни ACCOUNT_GMAIL (або залиш порожнім — спробує автодетект)
 *    3. Запусти  ▶ fullSetup()  — робить все одним кліком:
 *         - відмічає профіль як "Поданий Apeal" в таблиці Apeal
 *         - створює аркуш Mail Junk з форматуванням
 *         - ставить щогодинний тригер
 *    4. Повтори на кожному акаунті — дані йдуть в одну таблицю
 *
 *  Тестування:
 *    testSearchAppealEmails()  — пошук листів (лише лог, без запису)
 *    testWriteToSheet()        — пошук + запис у таблицю (тестово)
 *
 *  Аркуш "Mail Junk":
 *    Gmail    — пошта акаунту
 *    Mail     — sucess / cancel / reinstated / suspended
 *    Time     — час коли лист прийшов
 *    Status   — Відновленний / Відхилений
 *    Ban type — тип порушення з листа (напр. Circumventing Systems: Cloaking)
 * ================================================================
 */

// ─────────────────────────────────────────────────────────────────
//  НАЛАШТУВАННЯ
// ─────────────────────────────────────────────────────────────────

/** ID цільової таблиці */
const TARGET_SPREADSHEET_ID = '1z31pMXcXQXG6VRtgukbwjA7GCYHdR-qRL9HHF-wRK8E';

/**
 * Пошта акаунту. Рекомендовано заповнити вручну.
 * Якщо порожнє — спробує Session.getActiveUser() / getEffectiveUser().
 */
const ACCOUNT_GMAIL = '';

/**
 * Назва профілю Octo Browser (без тегів, напр. "GERM722").
 * Заповнюється автоматично при генерації скрипта через Landing Analyzer.
 * Записується в колонку F (Octo Name) кожного рядка.
 */
const OCTO_PROFILE_NAME = '';

/** Відправники листів Google Ads (масив — пошук по всіх) */
const ADS_SENDERS = [
  'ads-account-noreply@google.com',
  'ads-support@google.com',
  'google-ads-noreply@google.com',
  'adwords-noreply@google.com'
];

/** Максимум тредів за один запуск */
const MAX_THREADS = 50;

/** Скільки годин lookback при першому запуску */
const IMPORT_LOOKBACK_HOURS = 168;  // 7 днів

// ─────────────────────────────────────────────────────────────────
//  МАРКЕРИ АПЕЛЯЦІЙ
// ─────────────────────────────────────────────────────────────────

/**
 * Картинки в HTML листа — визначають тип результату апеляції.
 *   sucess.png → Відновленний
 *   cancel.png → Відхилений
 */
const APPEAL_IMG_APPROVED = 'sucess.png';
const APPEAL_IMG_REJECTED = 'cancel.png';

// ─────────────────────────────────────────────────────────────────
//  ТЕКСТОВІ МАРКЕРИ (англійська — перед перевіркою лист перекладається)
// ─────────────────────────────────────────────────────────────────

/**
 * Фрази-маркери в АНГЛІЙСЬКОМУ тексті (після перекладу через LanguageApp).
 * Порядок масиву = пріоритет перевірки.
 */
const TEXT_REINSTATED_PATTERNS = [
  'reinstated your account',
  'appeal has been approved',
  'appeal was approved',
  'account has been reinstated',
  'account was reinstated',
  'have reinstated your'
];

const TEXT_SUSPENDED_PATTERNS = [
  'account has been suspended',
  'account was suspended',
  'account is suspended',
  'have suspended your',
  'suspending your account'
];

/** Максимум символів для перекладу (економія квоти LanguageApp) */
const TRANSLATE_MAX_CHARS = 3000;

/**
 * Обов'язковий підпис для текстового деклайну.
 * Справжні деклайни підписані "gTech Customer Experience".
 * Листи-підтвердження ("ми отримали вашу апеляцію") підписані
 * "Google Ads Team" / "Zespół Google Ads" тощо — їх фільтруємо.
 */
const GTECH_SIGNATURE = 'gtech';

/**
 * Фрази-виключення (англійською, після перекладу).
 * Якщо лист містить будь-яку з них — це НЕ деклайн,
 * а лише підтвердження отримання апеляції.
 */
const TEXT_IGNORE_PATTERNS = [
  'we received your appeal',
  'we have received your appeal',
  'received the appeal',
  'thank you for submitting your appeal',
  'thank you for sending the appeal',
  'thank you for submitting an appeal',
  'your appeal is being reviewed',
  'wait for this appeal to be reviewed',
  'will review your appeal'
];

// ─────────────────────────────────────────────────────────────────
//  АРКУШ ТА ЗАГОЛОВКИ
// ─────────────────────────────────────────────────────────────────

const SHEET_NAME = 'Mail Junk';
const HEADERS    = ['Gmail', 'Mail', 'Time', 'Status', 'Ban type', 'Octo Name'];

// ─────────────────────────────────────────────────────────────────
//  SCRIPT PROPERTIES KEYS
// ─────────────────────────────────────────────────────────────────

const PROP_PROCESSED_IDS  = 'ADS_IMPORT_PROCESSED_IDS';
const PROP_IMPORT_START_MS = 'ADS_IMPORT_START_MS';

// ─────────────────────────────────────────────────────────────────
//  ▶▶▶ ТОЧКА ВХОДУ — запускати першим
// ─────────────────────────────────────────────────────────────────

/**
 * Запускає весь ланцюжок одноразово:
 *   1. Відмічає себе як "Поданий Apeal" в таблиці Apeal
 *   2. Запускає setupMailImporter() — створює аркуш, форматування
 *   3. Запускає createHourlyTrigger() — ставить щогодинний тригер
 */
function fullSetup() {
  Logger.log('[fullSetup] ▶ Старт повного налаштування');
  Logger.log('[fullSetup] Профіль: ' + (OCTO_PROFILE_NAME || '(не задано)'));

  // Крок 1: відмітити себе як поданий
  Logger.log('[fullSetup] ─── Крок 1: markSelfAsSubmitted');
  markSelfAsSubmitted();

  // Крок 2: setup аркушу Mail Junk
  Logger.log('[fullSetup] ─── Крок 2: setupMailImporter');
  setupMailImporter();

  // Крок 3: тригер щогодини
  Logger.log('[fullSetup] ─── Крок 3: createHourlyTrigger');
  createHourlyTrigger();

  Logger.log('[fullSetup] ✅ Готово! Всі 3 кроки виконано успішно.');
}

// ─────────────────────────────────────────────────────────────────
//  ГОЛОВНА ФУНКЦІЯ (тригер щогодини)
// ─────────────────────────────────────────────────────────────────

/**
 * Основна функція — викликається тригером щогодини.
 * Шукає нові листи з sucess.png / cancel.png і записує в "Mail Junk".
 */
function importAdsEmails() {
  const myEmail = resolveAccountEmail_();
  let ss;

  try {
    ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  } catch (e) {
    throw new Error('Не вдалось відкрити таблицю [' + TARGET_SPREADSHEET_ID + ']: ' + e.message);
  }

  try {
    const sheet = getOrCreateSheet_(ss);
    const props = PropertiesService.getScriptProperties();
    const processedIds = loadProcessedIds_(props);
    const importStartMs = getImportStartMs_(props);

    // Зберігаємо ім'я профілю у властивостях (для аудиту)
    if (OCTO_PROFILE_NAME) {
      props.setProperty('OCTO_PROFILE_NAME', OCTO_PROFILE_NAME);
    }

    const newMessages = fetchAppealMessages_(processedIds, importStartMs);

    if (newMessages.length === 0) {
      saveProcessedIds_(props, processedIds);
      Logger.log('[' + (OCTO_PROFILE_NAME || myEmail) + '] Нових апеляційних листів не знайдено.');
      return;
    }

    writeToSheet_(sheet, myEmail, newMessages);
    saveProcessedIds_(props, processedIds);
    SpreadsheetApp.flush();

    Logger.log('[' + (OCTO_PROFILE_NAME || myEmail) + '] Записано ' + newMessages.length + ' листів у "' + SHEET_NAME + '".');

  } catch (e) {
    Logger.log('ПОМИЛКА importAdsEmails: ' + e.message + '\n' + e.stack);
    throw e;
  }
}

// ─────────────────────────────────────────────────────────────────
//  SETUP & TRIGGER
// ─────────────────────────────────────────────────────────────────

/**
 * Одноразове налаштування — запусти перед тригером.
 * Створює аркуш, заголовки, форматування.
 */
function setupMailImporter() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const props = PropertiesService.getScriptProperties();

  const sheet = getOrCreateSheet_(ss);

  // Форматування
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, HEADERS.length);

  // Умовне форматування колонки Status (D)
  // Відновленний → зелений фон, Відхилений → червоний фон
  const statusRange = sheet.getRange('D2:D1000');
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Відновленний').setBackground('#b6d7a8').setFontColor('#1e4d2b')
      .setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Відхилений').setBackground('#ea9999').setFontColor('#660000')
      .setRanges([statusRange]).build()
  ];
  sheet.setConditionalFormatRules(rules);

  // Запам'ятати старт імпорту
  if (!props.getProperty(PROP_IMPORT_START_MS)) {
    props.setProperty(PROP_IMPORT_START_MS, String(Date.now()));
  }

  SpreadsheetApp.flush();
  Logger.log('Setup завершено. Аркуш "' + SHEET_NAME + '" готовий.');
  Logger.log('Octo Profile: ' + (OCTO_PROFILE_NAME || '(не задано)'));
  Logger.log('Тепер запусти createHourlyTrigger().');
}

/**
 * Створює щогодинний тригер. Видаляє попередній щоб не було дублів.
 */
function createHourlyTrigger() {
  const handler = 'importAdsEmails';
  const props = PropertiesService.getScriptProperties();

  // Видалити старі тригери
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === handler; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyHours(1)
    .create();

  props.setProperty(PROP_IMPORT_START_MS, String(Date.now()));

  Logger.log('Тригер створено: ' + handler + ' (щогодини).');
}

// ─────────────────────────────────────────────────────────────────
//  ТЕСТОВІ ФУНКЦІЇ
// ─────────────────────────────────────────────────────────────────

/**
 * ТЕСТ 1: Пошук апеляційних листів — лише логування, БЕЗ запису.
 * Ігнорує processedIds, шукає за останні IMPORT_LOOKBACK_HOURS годин.
 */
function testSearchAppealEmails() {
  const myEmail = resolveAccountEmail_();
  Logger.log('═══ ТЕСТ: Пошук апеляційних листів ═══');
  Logger.log('Акаунт: ' + myEmail);

  const startMs = Date.now() - (IMPORT_LOOKBACK_HOURS * 60 * 60 * 1000);
  const messages = fetchAppealMessages_(new Set(), startMs);

  Logger.log('Знайдено апеляційних листів (картинки + текст): ' + messages.length);

  if (messages.length === 0) {
    Logger.log('Жодного апеляційного листа не знайдено за останні ' + IMPORT_LOOKBACK_HOURS + ' годин.');
    Logger.log('Спробуй testSearchAllTime() щоб пошукати за весь час.');
    return;
  }

  messages.forEach(function (m, i) {
    Logger.log('#' + (i + 1) + ' | Mail: ' + m.mailType + ' | Method: ' + m.method + ' | Status: ' + m.status);
    Logger.log('  → Ban type: ' + (m.banType || '-'));
    Logger.log('  → Time: ' + m.dateReceived);
    Logger.log('  → Subject: ' + m.subject);
    Logger.log('  → MsgID: ' + m.msgId);
  });

  Logger.log('═══ Кінець тесту. Дані НЕ записані в таблицю. ═══');
}

/**
 * ТЕСТ 2: Пошук + ЗАПИС у таблицю — повний цикл тестово.
 * Ігнорує processedIds щоб записати все що знайде.
 */
function testWriteToSheet() {
  const myEmail = resolveAccountEmail_();
  Logger.log('═══ ТЕСТ: Пошук + запис у таблицю ═══');
  Logger.log('Акаунт: ' + myEmail);

  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(ss);

  // Шукаємо за ВЕСЬ час (без обмежень), щоб зловити навіть старі листи
  const messages = fetchAppealMessages_(new Set(), 0);

  Logger.log('Знайдено: ' + messages.length + ' листів');

  if (messages.length === 0) {
    Logger.log('Нічого не знайдено. Нічого не записано.');
    return;
  }

  writeToSheet_(sheet, myEmail, messages);
  SpreadsheetApp.flush();

  messages.forEach(function (m, i) {
    Logger.log(
      '#' + (i + 1) +
      ' | Mail: ' + m.mailType +
      ' | Method: ' + m.method +
      ' | Time: ' + m.dateReceived +
      ' | Status: ' + m.status
    );
  });

  Logger.log('═══ Записано ' + messages.length + ' рядків у "' + SHEET_NAME + '". ═══');
}

/**
 * ТЕСТ 3: Пошук за ВЕСЬ час (без lookback). Лише лог, без запису.
 */
function testSearchAllTime() {
  const myEmail = resolveAccountEmail_();
  Logger.log('═══ ТЕСТ: Пошук за весь час ═══');
  Logger.log('Акаунт: ' + myEmail);

  const messages = fetchAppealMessages_(new Set(), 0);

  Logger.log('Знайдено: ' + messages.length);
  messages.forEach(function (m, i) {
    Logger.log(
      '#' + (i + 1) +
      ' | Mail: ' + m.mailType +
      ' | Method: ' + m.method +
      ' | Time: ' + m.dateReceived +
      ' | Status: ' + m.status +
      ' | Subject: ' + m.subject
    );
  });

  Logger.log('═══ Кінець тесту. ═══');
}

/**
 * Допоміжний тест — перевірка email акаунту.
 */
function debugAccountEmail() {
  Logger.log('ACCOUNT_GMAIL const: "' + ACCOUNT_GMAIL + '"');
  Logger.log('Session.getActiveUser(): ' + safeGetEmail_(function () {
    return Session.getActiveUser().getEmail();
  }));
  Logger.log('Session.getEffectiveUser(): ' + safeGetEmail_(function () {
    return Session.getEffectiveUser().getEmail();
  }));
  Logger.log('Resolved → ' + resolveAccountEmail_());
}

// ─────────────────────────────────────────────────────────────────
//  УТИЛІТИ: RESET
// ─────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────
//  АВТО-СТАТУС "Поданий Apeal" в таблиці Apeal
// ─────────────────────────────────────────────────────────────────

/**
 * Одноразово знаходить себе в Apeal (кол. C = OCTO_PROFILE_NAME),
 * виставляє F = "Поданий Apeal" і показує popup.
 * Повторно не спрацьовує (захист через ScriptProperties).
 */
function markSelfAsSubmitted() {
  const PROP_MARKED = 'MARK_SELF_DONE_' + OCTO_PROFILE_NAME;
  const props = PropertiesService.getScriptProperties();

  if (!OCTO_PROFILE_NAME) {
    Logger.log('[❌ markSelfAsSubmitted] OCTO_PROFILE_NAME не задано — нічого не роблю.');
    return;
  }

  if (props.getProperty(PROP_MARKED) === 'true') {
    Logger.log('[ℹ️ markSelfAsSubmitted] Статус для "' + OCTO_PROFILE_NAME + '" вже було виставлено раніше. Для повтору запусти resetMarkSelf().');
    return;
  }

  const result = findSelfInApeal_();

  if (!result) {
    Logger.log('[⚠️ markSelfAsSubmitted] Профіль "' + OCTO_PROFILE_NAME + '" не знайдено в колонці C листа Apeal. Перевір назву.');
    return;
  }

  // Виставляємо F = "Поданий Apeal" + H = поточна дата + K = пошта акаунту
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const apealSheet = ss.getSheetByName('Apeal');
  const now = new Date();
  Logger.log('[markSelfAsSubmitted] Записуємо F рядок=' + result.row);
  apealSheet.getRange(result.row, 6).setValue('Поданий Apeal');
  Logger.log('[markSelfAsSubmitted] F записано. Записуємо H рядок=' + result.row + ' значення=' + now);
  const hCell = apealSheet.getRange(result.row, 8);
  hCell.setValue(now);
  hCell.setNumberFormat('dd.MM.yyyy HH:mm:ss');

  // Страховка: записуємо пошту акаунту в колонку K (11)
  const myEmail = resolveAccountEmail_();
  Logger.log('[markSelfAsSubmitted] Записуємо K (пошта) рядок=' + result.row + ' значення=' + myEmail);
  apealSheet.getRange(result.row, 11).setValue(myEmail);

  SpreadsheetApp.flush();
  Logger.log('[markSelfAsSubmitted] H після запису: ' + hCell.getValue() + ' | display: ' + hCell.getDisplayValue());
  Logger.log('[markSelfAsSubmitted] K після запису: ' + apealSheet.getRange(result.row, 11).getValue());

  props.setProperty(PROP_MARKED, 'true');

  Logger.log('[✅ markSelfAsSubmitted] Готово! Профіль="' + result.foundName + '" рядок=' + result.row + ' → F="Поданий Apeal" (повторне виконання заблоковано)');
  Logger.log('[markSelfAsSubmitted] ✅ Рядок ' + result.row + ' | C="' + result.foundName + '" → F="Поданий Apeal"');
}

/**
 * ТЕСТ — показує чи знайде себе в Apeal, БЕЗ запису.
 */
function testMarkSelf() {
  if (!OCTO_PROFILE_NAME) {
    Logger.log('[❌ testMarkSelf] OCTO_PROFILE_NAME не задано.');
    return;
  }

  const result = findSelfInApeal_();

  if (!result) {
    Logger.log('[🔍 testMarkSelf] Профіль "' + OCTO_PROFILE_NAME + '" НЕ знайдено в кол. C листа Apeal. Перевір: чи є такий рядок, чи немає зайвих пробілів.');
    return;
  }

  Logger.log(
    '[🔍 testMarkSelf] ТЕСТ (без запису)' +
    ' | Шукали: "' + OCTO_PROFILE_NAME + '"' +
    ' | Знайшли: "' + result.foundName + '" рядок=' + result.row +
    ' | F зараз: "' + result.currentStatus + '"' +
    ' | При реальному запуску → F="Поданий Apeal"'
  );
  Logger.log('[testMarkSelf] Знайдено: рядок=' + result.row + ' | C="' + result.foundName + '" | F="' + result.currentStatus + '"');
}

/**
 * Скидає захист — дозволяє повторно виконати markSelfAsSubmitted().
 */
function resetMarkSelf() {
  const PROP_MARKED = 'MARK_SELF_DONE_' + OCTO_PROFILE_NAME;
  PropertiesService.getScriptProperties().deleteProperty(PROP_MARKED);
  Logger.log('[resetMarkSelf] Захист скинуто для "' + OCTO_PROFILE_NAME + '"');
  Logger.log('[resetMarkSelf] 🔄 Захист скинуто для "' + OCTO_PROFILE_NAME + '". Тепер можна запустити markSelfAsSubmitted() знову.');
}

/**
 * Шукає OCTO_PROFILE_NAME в колонці C листа Apeal.
 * Ігнорує префікс [Transferred] (з пробілом і без).
 * @returns {{row, foundName, currentStatus} | null}
 */
function findSelfInApeal_() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const apealSheet = ss.getSheetByName('Apeal');
  if (!apealSheet) {
    Logger.log('[findSelfInApeal_] ❌ Лист Apeal не знайдено');
    return null;
  }

  const lastRow = apealSheet.getLastRow();
  if (lastRow < 2) return null;

  const data = apealSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const searchName = OCTO_PROFILE_NAME.trim().toLowerCase();

  for (let i = 0; i < data.length; i++) {
    const rawName = String(data[i][2] || '').trim(); // C
    // Видаляємо префікс [Transferred] (з пробілом або без)
    const cleanName = rawName.replace(/^\[Transferred\]\s*/i, '').trim().toLowerCase();

    if (cleanName === searchName) {
      return {
        row: i + 2,
        foundName: rawName,
        currentStatus: String(data[i][5] || '').trim() // F
      };
    }
  }

  return null;
}

// ─────────────────────────────────────────────────────────────────
//  УТИЛІТИ: RESET
// ─────────────────────────────────────────────────────────────────

/** Скидає processedIds — наступний запуск обробить заново. */
function resetProcessedIds() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_PROCESSED_IDS);
  Logger.log('Processed IDs скинуто.');
}

/** Скидає import start на зараз. */
function resetImportStartNow() {
  PropertiesService.getScriptProperties().setProperty(PROP_IMPORT_START_MS, String(Date.now()));
  Logger.log('Import start оновлено на поточний момент.');
}

// ─────────────────────────────────────────────────────────────────
//  ОСНОВНА ЛОГІКА
// ─────────────────────────────────────────────────────────────────

/**
 * Шукає в Gmail листи від ADS_SENDER, фільтрує тільки ті,
 * що містять sucess.png або cancel.png в HTML.
 *
 * @param {Set}    processedIds  — ID вже оброблених листів
 * @param {number} importStartMs — timestamp від якого шукати (0 = без обмежень)
 * @returns {Array<{msgId, dateReceived, subject, mailType, status}>}
 */
function fetchAppealMessages_(processedIds, importStartMs) {
  const query = buildSendersQuery_();
  let threads;

  const cutoffMs = importStartMs
    ? importStartMs - (IMPORT_LOOKBACK_HOURS * 60 * 60 * 1000)
    : 0;

  try {
    threads = GmailApp.search(query, 0, MAX_THREADS);
  } catch (e) {
    throw new Error('Gmail search failed: ' + e.message);
  }

  const results = [];

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var message = messages[m];
      var msgId = String(message.getId());
      var messageDate = message.getDate();

      // Пропускаємо старіші за cutoff
      if (cutoffMs && messageDate.getTime() < cutoffMs) continue;

      // Пропускаємо вже оброблені
      if (processedIds.has(msgId)) continue;

      // Парсимо HTML на наявність маркерних картинок
      var appealType = detectAppealType_(message);

      // Позначаємо як бачений
      processedIds.add(msgId);

      // Якщо лист не містить маркерних картинок — пропускаємо
      if (!appealType) continue;

      var subject = String(message.getSubject() || '');

      // Парсимо тип бану з HTML таблиці листа
      var banType = parseBanType_(message);

      results.push({
        msgId: msgId,
        dateReceived: messageDate,
        subject: subject,
        mailType: appealType.mailType,
        status: appealType.status,
        method: appealType.method || 'unknown',
        banType: banType
      });
    }
  }

  // Сортування: старіші → новіші
  results.sort(function (a, b) {
    return a.dateReceived.getTime() - b.dateReceived.getTime();
  });

  return results;
}

/**
 * Визначає тип апеляції двома методами:
 *   1) Маркерні картинки (sucess.png / cancel.png) — швидко, без API
 *   2) Текстовий аналіз з перекладом на англійську — LanguageApp
 * @returns {{mailType: string, status: string} | null}
 */
function detectAppealType_(message) {
  // 1. Спочатку — швидка перевірка по картинкам (без API)
  var imgResult = detectByImage_(message);
  if (imgResult) return imgResult;

  // 2. Якщо картинок нема — аналіз тексту з перекладом
  var textResult = detectByText_(message);
  if (textResult) return textResult;

  return null;
}

// ─── Детекція по картинкам ───────────────────────────────────────

function detectByImage_(message) {
  var html;
  try { html = String(message.getBody() || ''); } catch (_) { return null; }
  if (!html) return null;

  if (html.indexOf(APPEAL_IMG_APPROVED) !== -1) return { mailType: 'sucess',  status: 'Відновленний', method: 'image' };
  if (html.indexOf(APPEAL_IMG_REJECTED) !== -1) return { mailType: 'cancel',  status: 'Відхилений',   method: 'image' };
  return null;
}

// ─── Детекція по тексту (з перекладом) ───────────────────────────

/**
 * Бере plain text листа → перекладає на англійську →
 * шукає фрази-маркери suspended / reinstated.
 *
 * Для ДЕКЛАЙНУ (suspended) обов'язково:
 *   1) Лист підписаний "gTech" (Customer Experience)
 *   2) Лист НЕ є підтвердженням отримання апеляції
 */
function detectByText_(message) {
  var plainBody;
  try { plainBody = String(message.getPlainBody() || ''); } catch (_) { return null; }
  if (!plainBody || plainBody.length < 30) return null;

  var englishText = translateToEnglish_(plainBody);
  if (!englishText) return null;

  var lower = englishText.toLowerCase();

  // ─── Обов'язкова перевірка: текстовий детект ТІЛЬКИ з підписом gTech ───
  var hasGtech = lower.indexOf(GTECH_SIGNATURE) !== -1;
  if (!hasGtech) {
    Logger.log('⏭ Лист не містить "gTech" — текстовий детект пропущено.');
    return null;
  }

  // ─── Фільтр: пропускаємо листи-підтвердження отримання апеляції ───
  for (var i = 0; i < TEXT_IGNORE_PATTERNS.length; i++) {
    if (lower.indexOf(TEXT_IGNORE_PATTERNS[i]) !== -1) {
      Logger.log('⏭ Пропущено лист-підтвердження (ignore pattern: "' + TEXT_IGNORE_PATTERNS[i] + '")');
      return null;
    }
  }

  // ─── Reinstatement — більш специфічний, перевіряємо першим ───
  for (var i = 0; i < TEXT_REINSTATED_PATTERNS.length; i++) {
    if (lower.indexOf(TEXT_REINSTATED_PATTERNS[i]) !== -1) {
      return { mailType: 'reinstated', status: 'Відновленний', method: 'text+gTech' };
    }
  }

  // ─── Suspension ───
  for (var i = 0; i < TEXT_SUSPENDED_PATTERNS.length; i++) {
    if (lower.indexOf(TEXT_SUSPENDED_PATTERNS[i]) !== -1) {
      return { mailType: 'suspended', status: 'Відхилений', method: 'text+gTech' };
    }
  }

  return null;
}

/**
 * Перекладає текст на англійську через LanguageApp.
 * Авто-детект мови, обрізає до TRANSLATE_MAX_CHARS.
 * Якщо переклад не вдався — повертає оригінал (може бути англійською).
 */
function translateToEnglish_(text) {
  var chunk = text.substring(0, TRANSLATE_MAX_CHARS);

  // Швидка перевірка — якщо вже англійською, не витрачаємо квоту
  var lowerChunk = chunk.toLowerCase();
  for (var i = 0; i < TEXT_REINSTATED_PATTERNS.length; i++) {
    if (lowerChunk.indexOf(TEXT_REINSTATED_PATTERNS[i]) !== -1) return chunk;
  }
  for (var i = 0; i < TEXT_SUSPENDED_PATTERNS.length; i++) {
    if (lowerChunk.indexOf(TEXT_SUSPENDED_PATTERNS[i]) !== -1) return chunk;
  }

  // Спочатку пробуємо LanguageApp (якщо GCP прив'язано)
  try {
    return LanguageApp.translate(chunk, '', 'en');
  } catch (e) {
    Logger.log('⚠ LanguageApp недоступний: ' + e.message + ' → використовую fetch fallback');
  }

  // Fallback: Google Translate через публічний endpoint (без GCP проєкту)
  try {
    var encoded = encodeURIComponent(chunk);
    var url = 'https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=en&dt=t&q=' + encoded;
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('⚠ fetch translate HTTP ' + response.getResponseCode());
      return chunk;
    }
    var json = JSON.parse(response.getContentText());
    var translated = '';
    if (json && json[0]) {
      for (var i = 0; i < json[0].length; i++) {
        if (json[0][i] && json[0][i][0]) translated += json[0][i][0];
      }
    }
    return translated || chunk;
  } catch (e2) {
    Logger.log('⚠ fetch translate failed: ' + e2.message);
    return chunk;  // останній fallback — оригінал
  }
}

/**
 * Парсить тип бану (порушення) з HTML листа.
 * Шукає таблицю з колонками Type/Status, витягує текст типу порушення.
 * Приклад: "Circumventing Systems: Cloaking", "Malicious software" тощо.
 * @returns {string} тип бану або '' якщо не знайдено
 */
function parseBanType_(message) {
  var html;
  try { html = String(message.getBody() || ''); } catch (_) { return ''; }
  if (!html) return '';

  // Метод 1: Шукаємо <a> з текстом порушення перед "Unsuccessful" / "Successful"
  // Типова структура: <td><a href="...">Circumventing Systems: Cloaking</a></td> <td>...Unsuccessful</td>
  var banTypes = [];

  // Regex: знаходимо текст всередині <a> тегів в рядках таблиці
  // Google Ads листи мають лінки через c.gle або support.google.com
  // href може бути з лапками ("url") або без (url target=...)
  var linkRegex = /<a[^>]*href=(?:["']([^"']*(?:support\.google\.com|c\.gle)[^"']*)["']|([^\s>]*(?:support\.google\.com|c\.gle)[^\s>]*))[^>]*>([^<]+)<\/a>/gi;
  var match;
  while ((match = linkRegex.exec(html)) !== null) {
    var linkUrl = (match[1] || match[2] || '').toLowerCase();
    var linkText = (match[3] || '').trim();

    // Фільтр по URL — пропускаємо лінки дій (contact, answer, forms)
    if (linkUrl.indexOf('/contact/') !== -1 ||
        linkUrl.indexOf('/answer/') !== -1 ||
        linkUrl.indexOf('/forms/') !== -1) {
      continue;
    }

    // Фільтр по тексту — потрібні тільки назви порушень
    var lt = linkText.toLowerCase();
    if (linkText.length > 5 && linkText.length < 200 &&
        lt.indexOf('http') === -1 &&
        lt.indexOf('click') === -1 &&
        lt.indexOf('here') === -1 &&
        lt.indexOf('google') === -1 &&
        lt.indexOf('contact') === -1 &&
        lt.indexOf('help center') === -1 &&
        lt.indexOf('review') === -1 &&
        lt.indexOf('sign in') === -1 &&
        lt.indexOf('appeal') === -1 &&
        lt.indexOf('cancel your') === -1 &&
        lt.indexOf('request a') === -1 &&
        lt.indexOf('learn more') === -1 &&
        // Фільтр дій на різних мовах (укр, пол, нім, ісп, фр)
        lt.indexOf('оскаржити') === -1 &&
        lt.indexOf('скасувати') === -1 &&
        lt.indexOf('увійти') === -1 &&
        lt.indexOf('odwołaj') === -1 &&
        lt.indexOf('anuluj') === -1 &&
        lt.indexOf('zaloguj') === -1 &&
        lt.indexOf('einloggen') === -1 &&
        lt.indexOf('contestar') === -1 &&
        lt.indexOf('contester') === -1) {
      banTypes.push(linkText);
    }
  }

  // Метод 2: Шукаємо текст між <td> тегами перед Unsuccessful/Successful
  if (banTypes.length === 0) {
    var tdRegex = /<td[^>]*>\s*(?:<[^>]*>)*\s*([^<]{5,150})\s*(?:<[^>]*>)*\s*<\/td>\s*(?:<[^>]*>)*\s*<td[^>]*>[^<]*(?:Unsuccessful|Successful|unsuccessful|successful)/gi;
    while ((match = tdRegex.exec(html)) !== null) {
      var cellText = match[1].trim();
      if (cellText.length > 3 && cellText !== 'Type') {
        banTypes.push(cellText);
      }
    }
  }

  Logger.log('parseBanType_ Метод 1 (links): знайдено ' + banTypes.length + ' результатів');

  // Метод 3: gTech листи — бан тайп після <strong>Порушення правил</strong>
  // Структура: <strong>Порушення правил</strong><br>\n    Застосування неприйнятних бізнес-моделей</p>
  if (banTypes.length === 0) {
    // Варіант A: <strong>Header</strong><br> Text</p>
    var violationHeaderRegex = /<strong>([^<]*)<\/strong>\s*(?:<br\s*\/?>)?\s*([^<]{4,200})\s*<\/p>/gi;
    while ((match = violationHeaderRegex.exec(html)) !== null) {
      var headerText = match[1].trim();
      var bodyText = match[2].trim();
      var headerLower = headerText.toLowerCase();
      if (headerLower.indexOf('порушенн') !== -1 || headerLower.indexOf('violation') !== -1 ||
          headerLower.indexOf('policy') !== -1 || headerLower.indexOf('правил') !== -1 ||
          headerLower.indexOf('naruszeni') !== -1 || headerLower.indexOf('verstoß') !== -1 ||
          headerLower.indexOf('infracción') !== -1 || headerLower.indexOf('infraction') !== -1) {
        if (bodyText.length > 3 && bodyText.length < 200) {
          banTypes.push(bodyText);
        }
      }
    }
  }

  Logger.log('parseBanType_ Метод 3 (gTech HTML): знайдено ' + banTypes.length + ' результатів');

  // Метод 4: gTech plain text fallback — "*Порушення правил*\nТекст бану"
  if (banTypes.length === 0) {
    var plainBody;
    try { plainBody = String(message.getPlainBody() || ''); } catch (_) { plainBody = ''; }
    if (plainBody) {
      var plainViolationRegex = /\*([^*]*(?:порушенн|violation|policy|правил|naruszeni|verstoß|infracción|infraction)[^*]*)\*\s*\n\s*(.+)/gi;
      while ((match = plainViolationRegex.exec(plainBody)) !== null) {
        var pBodyText = match[2].trim();
        if (pBodyText.length > 3 && pBodyText.length < 200) {
          banTypes.push(pBodyText);
        }
      }
    }
  }

  if (banTypes.length === 0) return '';

  // Видаляємо дублі та повертаємо через " | "
  var unique = [];
  var seen = {};
  for (var i = 0; i < banTypes.length; i++) {
    var key = banTypes[i].toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      unique.push(banTypes[i]);
    }
  }

  return unique.join(' | ');
}

/**
 * Записує знайдені листи в аркуш "Mail Junk".
 * Кожен лист = 1 рядок: Gmail | Mail | Time | Status | Ban type
 */
function writeToSheet_(sheet, myEmail, messages) {
  var startRow = sheet.getLastRow() + 1;

  var rows = messages.map(function (m) {
    return [
      myEmail,              // Gmail
      m.mailType,           // Mail  (sucess / cancel)
      m.dateReceived,       // Time  (коли лист прийшов)
      m.status,             // Status (Відновленний / Відхилений)
      m.banType || '',      // Ban type (тип порушення)
      OCTO_PROFILE_NAME     // Octo Name (профіль з якого відправлено апеляцію)
    ];
  });

  sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
}

// ─────────────────────────────────────────────────────────────────
//  ДОПОМІЖНІ ФУНКЦІЇ
// ─────────────────────────────────────────────────────────────────

/** Будує Gmail-запит по масиву відправників. */
function buildSendersQuery_() {
  if (ADS_SENDERS.length === 1) return 'from:' + ADS_SENDERS[0];
  return 'from:(' + ADS_SENDERS.join(' OR ') + ')';
}

/** Повертає (або створює) аркуш "Mail Junk" з заголовками. */
function getOrCreateSheet_(ss) {
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  var first = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var isEmpty = first.every(function (v) { return String(v).trim() === ''; });
  if (isEmpty) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  return sh;
}

/** Визначає email акаунту. */
function resolveAccountEmail_() {
  var manualEmail = String(ACCOUNT_GMAIL || '').trim().toLowerCase();
  if (manualEmail) return manualEmail;

  var activeEmail = safeGetEmail_(function () {
    return Session.getActiveUser().getEmail();
  });
  if (activeEmail) return activeEmail;

  var effectiveEmail = safeGetEmail_(function () {
    return Session.getEffectiveUser().getEmail();
  });
  if (effectiveEmail) return effectiveEmail;

  throw new Error(
    'Не вдалося визначити Gmail. Заповни const ACCOUNT_GMAIL вручну.'
  );
}

function safeGetEmail_(getter) {
  try {
    return String(getter() || '').trim().toLowerCase();
  } catch (_) {
    return '';
  }
}

function getImportStartMs_(props) {
  var value = Number(props.getProperty(PROP_IMPORT_START_MS) || 0);
  if (value > 0) return value;

  var now = Date.now();
  props.setProperty(PROP_IMPORT_START_MS, String(now));
  return now;
}

function loadProcessedIds_(props) {
  try {
    var raw = props.getProperty(PROP_PROCESSED_IDS) || '[]';
    return new Set(JSON.parse(raw).map(String));
  } catch (_) {
    return new Set();
  }
}

function saveProcessedIds_(props, setObj) {
  var arr = Array.from(setObj);
  var trimmed = arr.slice(Math.max(0, arr.length - 3000));
  props.setProperty(PROP_PROCESSED_IDS, JSON.stringify(trimmed));
}
