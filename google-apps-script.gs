/**
 * Скрипт для приёма анкет с сайта Aru Design в Google Таблицу.
 *
 * Как подключить (один раз, ~5 минут):
 * 1. Откройте sheets.google.com и создайте новую таблицу, назовите её, например, «Анкеты Aru Design».
 * 2. В меню таблицы: Расширения → Apps Script.
 * 3. Удалите всё в открывшемся редакторе и вставьте целиком этот файл. Нажмите 💾 (Сохранить).
 * 4. Справа вверху: «Начать развертывание» → «Новое развертывание».
 *    - Тип: «Веб-приложение»
 *    - «Выполнять от имени»: От моего имени
 *    - «У кого есть доступ»: Все
 *    Нажмите «Начать развертывание» и разрешите доступ (Google предупредит — это нормально,
 *    скрипт ваш собственный и пишет только в вашу таблицу).
 * 5. Скопируйте выданный «URL веб-приложения» (заканчивается на /exec)
 *    и вставьте его в index.html в строку SHEET_URL = "…".
 *
 * Каждая отправленная анкета появится в таблице новой строкой.
 * Первая строка (заголовки колонок) создастся автоматически при первой анкете.
 */

var SHEET_NAME = "Анкеты";

/**
 * Диагностика: откройте URL веб-приложения (/exec) в браузере —
 * увидите название и ссылку таблицы, куда пишутся анкеты, и число записей.
 */
function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss ? ss.getSheetByName(SHEET_NAME) : null;
  var info = {
    "куда_пишутся_анкеты": ss ? ss.getName() : "ОШИБКА: скрипт не привязан к таблице",
    "ссылка_на_таблицу": ss ? ss.getUrl() : "создайте скрипт через меню таблицы: Расширения → Apps Script",
    "вкладка": sh ? SHEET_NAME : "вкладка «" + SHEET_NAME + "» ещё не создана (не было ни одной анкеты)",
    "записей_анкет": sh ? Math.max(sh.getLastRow() - 1, 0) : 0
  };
  return ContentService.createTextOutput(JSON.stringify(info, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); // защита от одновременной записи двух анкет
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    var data = JSON.parse(e.postData.contents);
    var keys = Object.keys(data);

    // первая анкета — создаём строку заголовков
    if (sh.getLastRow() === 0) {
      sh.appendRow(keys);
      sh.getRange(1, 1, 1, keys.length).setFontWeight("bold");
      sh.setFrozenRows(1);
    }

    // если появились новые вопросы — добавляем недостающие колонки
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    keys.forEach(function (k) {
      if (headers.indexOf(k) === -1) {
        sh.getRange(1, headers.length + 1).setValue(k).setFontWeight("bold");
        headers.push(k);
      }
    });

    // записываем ответы в свои колонки
    sh.appendRow(headers.map(function (h) { return data[h] || ""; }));

    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}
