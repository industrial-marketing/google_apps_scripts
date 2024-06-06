
var functionsMap = {
    'generateSalesReportCommand': generateSalesReport,
    'generateBenchReport': generateBenchReport,
    'generatePaidHoursReport': generateWeekReport,
    'generatePaidHoursReportLastWeek': generateWeekReportLastWeek,
    'generateAllReport': generateAllReport,
    'generateAllReportLastWeek': generateAllReportLastWeek,
    'gatherDataInSheet': gatherDataInSheet,
    'gatherDataInSheetLastWeek': gatherDataInSheetLastWeek,
    'generateGatherScrumFilesDataFromFolder': gatherScrumFilesDataFromFolder,
    // 'copyDataToCompetencesSheet': copyDataToCompetencesSheet,
    'updateDeveloperStackData': updateDeveloperStackData,
    'collectCandidatesData': collectCandidatesData,
    'collectDeveloperCvData': collectDeveloperCvData,
    'collectDeveloperUpworkData': collectDeveloperUpworkData,
    'collectDeveloperVacationData': collectDeveloperVacationData
};


var monthNamesShort = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

function saveOAuthToken() {
    var token = ScriptApp.getOAuthToken();
    PropertiesService.getScriptProperties().setProperty('OAUTH_TOKEN', token);
}

// function saveOAuthSlackToken() {
//     var token = getOAuthService();
//     PropertiesService.getScriptProperties().setProperty('OAUTH_TOKEN', token);
// }

function getOAuthService() {
    return OAuth2.createService('Slack')
        .setAuthorizationBaseUrl('https://slack.com/oauth/v2/authorize')
        .setTokenUrl('https://slack.com/api/oauth.v2.access')
        .setClientId('6690155511488.7180362243282')  // Replace with your Slack app's client ID
        .setClientSecret('f49bebba3470b85defca2ea307ebc2e8')  // Replace with your Slack app's client secret
        .setCallbackFunction('authCallback')
        .setPropertyStore(PropertiesService.getUserProperties())
        .setScope('chat:write')
        .setScope('users:read')
        .setScope('users:read.email')
        .setParam('access_type', 'offline')
        .setParam('prompt', 'consent')
        .setParam('user_scope', 'chat:write');
}

function authCallback(request) {
    var service = getOAuthService();
    var authorized = service.handleCallback(request);
    if (authorized) {
        return HtmlService.createHtmlOutput('Success! You can close this tab.');
    } else {
        return HtmlService.createHtmlOutput('Denied. You can close this tab');
    }
}

function getAuthorizationUrl() {
    var service = getOAuthService();
    return service.getAuthorizationUrl();
}

function authorize() {
    var url = getAuthorizationUrl();
    var htmlOutput = HtmlService.createHtmlOutput('<a href="' + url + '" target="_blank">Authorize</a>');
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Authorize Slack');
}


function logout() {
    getOAuthService().reset();
}

function isAuth() {
    return getOAuthService().hasAccess();
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Обновление')
        .addItem('Получить токен', 'saveOAuthToken')
        .addItem('Обновить SALES report', 'wrapperGenerateSalesReport')
        .addItem('Обновить BENCH report', 'wrapperGenerateBenchReport')
        .addItem('Обновить ALL report', 'wrapperGenerateAllReport')
        .addItem('Обновить ALL report last week', 'wrapperGenerateAllReportLastWeek')
        .addItem('Обновить PAID HOURS report', 'wrapperGeneratePaidHoursReport')
        .addItem('Обновить PAID HOURS report last week', 'wrapperGeneratePaidHoursReportLastWeek')
        .addItem('Обновить Keywords', 'gatherKeywords')
        .addItem('Обновить Scrum files for current week', 'wrapperGatherDataInSheet')
        .addItem('Обновить Scrum files for last week', 'wrapperGatherDataInSheetLastWeek')
        .addItem('Обновить Scrum files 2024', 'wrapperGenerateGatherScrumFilesDataFromFolder')
        .addItem('Обновить DeveloperStackData', 'wrapperUpdateDeveloperStackData')
        .addItem('Обновить DeveloperCVData', 'wrapperCollectDeveloperCvData')
        .addItem('Обновить DeveloperCandidatesData', 'wrapperCollectCandidatesData')
        .addItem('Обновить DeveloperUpworkData', 'wrapperCollectDeveloperUpworkData')
        .addItem('Обновить DeveloperVacationData', 'wrapperCollectDeveloperVacationData')
        .addItem('Slack авторизация', 'authorize')
        .addToUi();
    ui.createMenu('Фильтры')
        .addItem('Показать все строки', 'showAllRows')
        .addItem('Показать все колонки', 'showAllColumns')
        .addItem('Скрыть текстовые колонки','hideTextInfo')
        .addItem('Только бенч', 'showOnlyBenchRows')
        .addItem('Выбор стеков', 'showStacksDialog')
        .addItem('Поиск', 'showSearchDialog')
        // .addItem('Сортировать по A-Z', 'sortDataAscending')
        // .addItem('Сортировать по Z-A', 'sortDataDescending')
        .addToUi();
}


function wrapperGenerateSalesReport() {
    runFunctionWithOAuthCheck('generateSalesReportCommand');
}
function generateSalesReportCommand() {
    generateSalesReport();
}

function wrapperGeneratePaidHoursReport() {
    runFunctionWithOAuthCheck('generatePaidHoursReport');
}
function generatePaidHoursReport() {
    generateWeekReport(false);
}

function wrapperGeneratePaidHoursReportLastWeek() {
    runFunctionWithOAuthCheck('generatePaidHoursReportLastWeek');
}
function generatePaidHoursReportLastWeek() {
    generateWeekReportLastWeek();
}

function wrapperGenerateGatherScrumFilesDataFromFolder() {
    runFunctionWithOAuthCheck('generateGatherScrumFilesDataFromFolder');
}
function generateGatherScrumFilesDataFromFolder() {
    gatherScrumFilesDataFromFolder();
}

function wrapperGenerateBenchReport() {
    runFunctionWithOAuthCheck('generateBenchReport');
}
function generateBenchReport() {
    generateSalesReport(false, false, true);
}


function wrapperGenerateAllReport() {
    runFunctionWithOAuthCheck('generateAllReport');
}
function generateAllReport() {
    generateSalesReport(true);
}


function wrapperGenerateAllReportLastWeek() {
    runFunctionWithOAuthCheck('generateAllReportLastWeek');
}
function generateAllReportLastWeek() {
    generateSalesReport(true, true);
}


function wrapperGatherDataInSheet() {
    runFunctionWithOAuthCheck('gatherDataInSheet');
}


function wrapperGatherDataInSheetLastWeek() {
    runFunctionWithOAuthCheck('gatherDataInSheetLastWeek');
}
function gatherDataInSheetLastWeek() {
    gatherDataInSheet(true);
}


// function wrapperCopyDataToCompetencesSheet() {
//     runFunctionWithOAuthCheck('copyDataToCompetencesSheet');
// }


function wrapperUpdateDeveloperStackData() {
    runFunctionWithOAuthCheck('updateDeveloperStackData');
}

function wrapperCollectCandidatesData() {
    runFunctionWithOAuthCheck('collectCandidatesData');
}

function wrapperCollectDeveloperCvData() {
    runFunctionWithOAuthCheck('collectDeveloperCvData');
}

function wrapperCollectDeveloperUpworkData() {
    runFunctionWithOAuthCheck('collectDeveloperUpworkData');
}

function wrapperCollectDeveloperVacationData() {
    runFunctionWithOAuthCheck('collectDeveloperVacationData');
}


function runFunctionWithOAuthCheck(functionName) {
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию с именем functionName
        functionsMap[functionName]();
    } else {
        // Токен отсутствует, отображаем диалоговое окно пользователю
        var response = Browser.msgBox(
            "OAuth Token Required",
            "Please obtain an OAuth token by clicking the 'OK' button.",
            Browser.Buttons.OK_CANCEL
        );

        if (response === Browser.Buttons.OK) {
            // Пользователь нажал OK, выполняем действия для получения токена
            saveOAuthToken();
            // После получения токена, повторяем попытку выполнения функции
            functionsMap[functionName](...params);
        }
        // Пользователь нажал Cancel, не выполняем функцию
    }
}

function showStacksDialog() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Stacks.html')
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Стеки');
}

function showProjectsDialog() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Projects.html')
        .setWidth(400)
        .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Проекты');
}

function showSearchDialog() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Search.html')
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Поиск');
}


function sortDataAscending() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var sortColumn = sheet.getActiveCell().getColumn();

    var dataRange = sheet.getRange(7, 1, sheet.getLastRow() - 6, sheet.getLastColumn());

    dataRange.sort([{column: sortColumn, ascending: true}]);
}

function sortDataDescending() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var sortColumn = sheet.getActiveCell().getColumn();

    var dataRange = sheet.getRange(7, 1, sheet.getLastRow() - 6, sheet.getLastColumn());

    dataRange.sort([{column: sortColumn, ascending: false}]);
}

function getStacks() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var stackRow = sheet.getRange("5:5");
    var values = stackRow.getValues();

    let stacks = [];
    for (let stack of values[0].slice(10)) {
        if (stack === 'Plan') {
            break;
        }
        stacks.push(stack);
    }

    return stacks;
}

function rgbToHex(rgb) {
    let match = rgb.match(/^rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i);
    if (!match) return '#ffffff';
    return "#" + ((1 << 24) + (+match[1] << 16) + (+match[2] << 8) + +match[3]).toString(16).slice(1).toUpperCase();
}

function getActiveStacks() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    // Logger.log(sheetName);

    var headerValues = sheet.getRange("5:5").getValues()[0];
    var startColumn = 11;
    var endColumn = headerValues.indexOf('Plan');

    var activeStacks = [];

    // Проверка, есть ли фильтр
    var filter = sheet.getFilter();
    if (!filter) {
        return activeStacks;
    }

    for (var i = startColumn; i <= endColumn; i++) {
        var criteria = filter.getColumnFilterCriteria(i);
        if (criteria && criteria.getCriteriaType()) {
            activeStacks.push(headerValues[i - 1]);  // Помните, что индексы в JavaScript начинаются с 0
        }
    }

    // Logger.log(activeStacks); // Выводим activeStacks в Logger
    return activeStacks;
}


function enableStack(stackName) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();
    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var headerRow = sheet.getRange("5:5");
    var headerValues = headerRow.getValues()[0];
    var columnIndex = headerValues.indexOf(stackName) + 1;

    var range = sheet.getFilter() || sheet.getRange(6, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    if (range) range.createFilter();
    var criteria = SpreadsheetApp.newFilterCriteria()
        .whenCellNotEmpty()
        .build();

    filter.setColumnFilterCriteria(columnIndex, criteria);
}

function disableStack(stackName) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();
    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var headerRow = sheet.getRange("5:5");
    var headerValues = headerRow.getValues()[0];
    var columnIndex = headerValues.indexOf(stackName) + 1;

    var filter = sheet.getFilter();
    if (filter) {
        filter.removeColumnFilterCriteria(columnIndex);
    }
}


function hideEmptyRows(sheet, sortColumn) {
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var dataRange = sheet.getRange(7, sortColumn, sheet.getLastRow() - 6);
    var data = dataRange.getValues();

    for (var i = 0; i < data.length; i++) {
        if (data[i][0] === "" || data[i][0] === null) {
            sheet.hideRows(i + 7); // Скрыть пустые строки
        }
    }
}

function searchData(query) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();
    var lastRow = sheet.getLastRow();

    // Определяем колонку для поиска и колонки для отображения запроса в зависимости от листа
    var columnToSearch, queryDisplayColumnStart, queryDisplayColumnEnd;
    switch(sheetName) {
        case "SALES report":
            columnToSearch = "M";
            queryDisplayColumnStart = "L";
            queryDisplayColumnEnd = "M";
            break;
        case "ALL report":
            columnToSearch = "M";
            queryDisplayColumnStart = "L";
            queryDisplayColumnEnd = "M";
            break;
        case "ALL report last week":
            columnToSearch = "I";
            queryDisplayColumnStart = "H";
            queryDisplayColumnEnd = "I";
            break;
        default:
            // Если лист не подходит под критерии, выходим из функции
            SpreadsheetApp.getUi().alert('Этот лист не поддерживает поиск по заметкам.');
            return;
    }

    var notesRange = sheet.getRange(columnToSearch + "7:" + columnToSearch + lastRow); // Получаем заметки начиная с 7-й строки до последней
    var notes = notesRange.getNotes();

    // Отображение запроса в указанных колонках
    if (query) {
        sheet.getRange(queryDisplayColumnStart + "3").setValue("Поиск по запросу:").setBackground('black').setFontColor('white');
        sheet.getRange(queryDisplayColumnEnd + "3").setValue(query).setBackground('black').setFontColor('white');
    } else {
        sheet.getRange(queryDisplayColumnStart + "3").setValue("").setBackground('white').setFontColor('black');
        sheet.getRange(queryDisplayColumnEnd + "3").setValue("").setBackground('white').setFontColor('black');
    }

    // Скрываем все строки начиная с 7-й
    sheet.hideRows(7, lastRow - 6);

    // Преобразуем запрос в массив слов в нижнем регистре
    var queryWords = query.toLowerCase().split(" ");

    // Идем по всем заметкам и показываем строки, которые содержат все слова из запроса
    notes.forEach(function(note, index) {
        var noteContent = note[0].toLowerCase();
        var containsAllWords = queryWords.every(function(word) {
            return noteContent.includes(word);
        });

        if (containsAllWords) {
            // Показываем строку, соответствующую найденной заметке
            sheet.showRows(index + 7);
        }
    });
}

function showAllRows() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    // Проверяем, есть ли скрытые строки
    var maxRows = sheet.getLastRow();
    var hiddenRows = [];
    for (var i = 1; i <= maxRows; i++) {
        if (sheet.isRowHiddenByUser(i)) {
            hiddenRows.push(i);
        }
    }

    // Если есть скрытые строки, показываем их
    if (hiddenRows.length > 0) {
        sheet.showRows(1, maxRows);
    }

    // Если есть активный фильтр, удаляем его
    var filter = sheet.getFilter();
    if (filter) {
        filter.remove();
    }

    // создаем новый фильтр
    // var headerRow = sheet.getRange("5:5");
    // var headerValues = headerRow.getValues()[0];
    // var endColumn = headerValues.indexOf('Plan');
    // var range = sheet.getRange(6, 2, sheet.getLastRow() - 5, endColumn-10);
    // range.createFilter();

    var lastColumn = sheet.getLastColumn();
    var range = sheet.getRange(6, 2, sheet.getLastRow() - 5, lastColumn);
    if (range) range.createFilter();

    // Сбрасываем поиск

    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();
    var lastRow = sheet.getLastRow();

    // Определяем колонку для поиска и колонки для отображения запроса в зависимости от листа
    var columnToSearch, queryDisplayColumnStart, queryDisplayColumnEnd;

    switch(sheetName) {
        case "SALES report":
            queryDisplayColumnStart = "L";
            queryDisplayColumnEnd = "M";
            break;
        case "ALL report":
            queryDisplayColumnStart = "L";
            queryDisplayColumnEnd = "M";
            break;
        case "ALL report last week":
            queryDisplayColumnStart = "H";
            queryDisplayColumnEnd = "I";
            break;
        default:
            // Если лист не подходит под критерии, выходим из функции
            return;
    }

    sheet.getRange(queryDisplayColumnStart + "3:" + queryDisplayColumnEnd + "3").clearContent().setBackground('white').setFontColor('black');

    // Определите диапазон строк, которому вы хотите задать новую высоту
    sheet.setRowHeights(1, 1, 1);
    sheet.setRowHeights(2, 1, 20);
    sheet.setRowHeights(3, 1, 40);
    sheet.setRowHeights(4, 1, 20);
    sheet.setRowHeights(5, 1, 80);
    sheet.setRowHeights(6, 1, 20);
    var startRow = 7;
    var numRows = sheet.getLastRow() - startRow + 1;
    // Установите высоту всех строк в этом диапазоне
    sheet.setRowHeights(startRow, numRows, 50);

    return true;
}

function hideTextInfo() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();
    var validSheetNames = ['ALL report', 'SALES report', 'ALL report last week'];

    if (validSheetNames.indexOf(sheetName) === -1) {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var headersRow = 5;
    var headersToHide = ["CV", "Upwork", "Stack", "Plan", "Fact"];

    var headersRange = sheet.getRange(headersRow, 1, 1, sheet.getLastColumn());
    var headersValues = headersRange.getValues()[0];

    for (var i = 0; i < headersValues.length; i++) {
        var header = headersValues[i];
        if (headersToHide.indexOf(header) !== -1) {
            sheet.hideColumns(i + 1);
        }
    }
}

function showAllColumns() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();
    var validSheetNames = ['ALL report', 'SALES report', 'ALL report last week'];

    if (validSheetNames.indexOf(sheetName) === -1) {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var headersRow = 5;
    var headersToHide = ["CV", "Upwork", "Stack", "Plan", "Fact"];

    var headersRange = sheet.getRange(headersRow, 1, 1, sheet.getLastColumn());
    var headersValues = headersRange.getValues()[0];

    for (var i = 0; i < headersValues.length; i++) {
        var header = headersValues[i];
        if (headersToHide.indexOf(header) !== -1) {
            sheet.showColumns(i + 1);
        }
    }
}


function getCurrentSearchQuery() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getSheetName();
    switch(sheetName) {
        case "SALES report":
            queryDisplayColumn = "M";
            break;
        case "ALL report":
            queryDisplayColumn = "L";
            break;
        case "ALL report last week":
            queryDisplayColumn = "I";
            break;
        default:
            // Если лист не подходит под критерии, выходим из функции
            return;
    }
    var query = sheet.getRange(queryDisplayColumn + "3").getValue();
    return query;
}

function getKeywords() {
    var keywordSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Keywords');
    if (!keywordSheet) {
        return [];  // Если нет листа "Keywords", возвращаем пустой массив
    }

    // Получаем все слова из листа "Keywords"
    var range = keywordSheet.getRange(1, 1, keywordSheet.getLastRow(), 1);
    var values = range.getValues();
    // values = [''];
    // Преобразуем двумерный массив в одномерный
    var keywords = [].concat(...values);

    //keywords = [];
    return keywords;
}



function showOnlyBenchRows() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var data = sheet.getDataRange().getValues();

    sheet.showRows(1, data.length); // Обязательно показываем все строки перед скрытием

    for (var i = 6; i < data.length; i++) {
        // Проверяем колонки E (индекс 4 соответственно)
        if (data[i][4] < 10) {
            sheet.hideRows(i + 1);
        }
    }


    // Сбрасываем поиск
    switch(sheetName) {
        case "SALES report":
            queryDisplayColumnStart = "L";
            queryDisplayColumnEnd = "M";
            break;
        case "ALL report":
            queryDisplayColumnStart = "L";
            queryDisplayColumnEnd = "M";
            break;
        case "ALL report last week":
            queryDisplayColumnStart = "H";
            queryDisplayColumnEnd = "I";
            break;
        default:
            // Если лист не подходит под критерии, выходим из функции
            return;
    }

    sheet.getRange(queryDisplayColumnStart + "3:" + queryDisplayColumnEnd + "3").clearContent().setBackground('white').setFontColor('black');

}



function checkOAuthToken() {
    var token = PropertiesService.getScriptProperties().getProperty('OAUTH_TOKEN');
    if (!token) {
        return false; // Если токен отсутствует, считаем его недействительным
    }

    var documentId = "1--eHqlntnVnOGlz_7mcpiXzu2YMN49P7y66yt2fpb6g"; // проверка сделана на рандомном id - SIAMEN AUSIANIKAU (Noda js+Angular)
    var url = "https://docs.googleapis.com/v1/documents/" + documentId; // URL для проверки токена

    var options = {
        "headers": {
            "Authorization": "Bearer " + token
        },
        "muteHttpExceptions": true // Включаем параметр для получения полного ответа сервера при возникновении ошибки
    };

    try {
        var response = UrlFetchApp.fetch(url, options);
        var statusCode = response.getResponseCode();

        if (statusCode === 200) {
            return true; // Токен действителен
        } else {
            return false; // Токен недействителен или возникла другая ошибка
        }
    } catch (e) {
        return false; // Произошла ошибка при запросе к API
    }
}

function generateSalesReport(all = false, isLastWeek = false, isBench = false) {


    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const benchSheetId = "1fpe40DxU-diKV_MfQayPIsBlTGDPBWeBrifCyUGdhy4";
    const benchSpreadsheet = SpreadsheetApp.openById(benchSheetId);
    const developerVacationSheet = ss.getSheetByName("DeveloperVacation");
    const developerVacations = developerVacationSheet.getDataRange().getValues();
    const developerRateSheet = ss.getSheetByName("DeveloperRates");
    const developerRates = developerRateSheet.getDataRange().getValues();

    const developersSheetId = "1VW615PcoaR90HLDD-JQeDmeAcz6DH1T_gCuN17v9C1I";
    const developersSheet = SpreadsheetApp.openById(developersSheetId);
    const developersPortfolioSheet = developersSheet.getSheetByName("DeveloperProjectData");
    const developerProfileSheet = developersSheet.getSheetByName("DeveloperProfiles");
    const developerProfiles = developerProfileSheet.getDataRange().getValues();
    const developerNameSheet = developersSheet.getSheetByName("Developers english vs russian names");
    const developerNames = developerNameSheet.getDataRange().getValues();

    let reportName = "SALES report"

    if(all)
        reportName = 'ALL report';
    if(isBench)
        reportName = 'SharpDev Bench Report';

    // Если флаг isLastWeek установлен в true, добавляем 'last week' к имени отчета
    if(isLastWeek) {
        reportName += ' last week';
    }

    // Проверяем, запущена ли функция на правильном листе
    if (reportName !== 'SharpDev Bench Report' && reportName !== 'ALL report' && reportName !== 'SALES report' && reportName !== 'ALL report last week') {
        // Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    const workloadSheetId = "1N65NUtqBA855C6K8swmeFQ9HbvIZU4fq4EnhYzvNV7Q";
    const workloadSpreadsheet = SpreadsheetApp.openById(workloadSheetId);

    let mondayDate, sundayDate;

    if (isLastWeek) {
        mondayDate = getLastMonday();
        sundayDate = getLastSunday();
    } else {
        mondayDate = getCurrentMonday();
        sundayDate = getCurrentSunday();
    }

    const mondayString = Utilities.formatDate(mondayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();
    const sundayString = Utilities.formatDate(sundayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();

    let workloadSheetName = mondayDate.getMonth() === sundayDate.getMonth() ?
        `${mondayString.split(" ")[0]}-${sundayString.split(" ")[0]} ${sundayString.split(" ")[1]}` :
        `${mondayString}-${sundayString}`;

    const workloadSheet = workloadSpreadsheet.getSheetByName(workloadSheetName);
    if (!workloadSheet) {
        //SpreadsheetApp.getUi().alert(`Cannot find sheet "${workloadSheetName}" in the workload spreadsheet.`);
        return;
    }

    let scrumSheetName = isLastWeek ? 'Scrum files for last week' : 'Scrum files for current week';
    const scrumSheet = ss.getSheetByName(scrumSheetName);

    // Get data from column E
    let columnEData = scrumSheet.getRange("E:E").getValues();

    // Sum all the values in column E
    let sumColumnE = columnEData.reduce((sum, value) => {
        // Ensure that the value is a number before adding it to the sum
        if (typeof value[0] == "number") {
            return sum + value[0];
        }
        return sum;
    }, 0);

    let developers = getDevelopers(workloadSheet, all, null, null, isLastWeek);
    let projectsData = getProjects(workloadSheet, null, isLastWeek);
    let projects = projectsData.projects;

    let projectsNextWeekData = getProjects(null, null, false, false, 1, workloadSpreadsheet);
    let projectsNextWeek = projectsNextWeekData.projects;
    let workloadNextWeekSheetName = projectsNextWeekData.workloadSheetName;

    let projectsNextNextWeekData = getProjects(null, null, false, false, 2, workloadSpreadsheet);
    let projectsNextNextWeek = projectsNextNextWeekData.projects;
    let workloadNextNextWeekSheetName = projectsNextNextWeekData.workloadSheetName;

    let projectsNextNextNextWeekData = getProjects(null, null, false, false, 3, workloadSpreadsheet);
    let projectsNextNextNextWeek = projectsNextNextNextWeekData.projects;
    let workloadNextNextNextWeekSheetName = projectsNextNextNextWeekData.workloadSheetName;

    let nextWeekProjects = {
        week: projects,
        nextWeek: projectsNextWeek,
        nextNextWeek: projectsNextNextWeek,
        nextNextNextWeek: projectsNextNextNextWeek,
    };


    // Считываем данные без заголовков
    const data = developersPortfolioSheet.getRange(2, 1, developersPortfolioSheet.getLastRow() - 1, developersPortfolioSheet.getLastColumn()).getValues();

    const developersProjects = {};

    // Структурирование данных по именам разработчиков
    data.forEach(row => {
        const developerName = row[0]; // Предполагаем, что имя разработчика находится в первой колонке
        if (!developersProjects[developerName]) {
            developersProjects[developerName] = [];
        }
        developersProjects[developerName].push(row.slice(1)); // Добавляем информацию о проекте, исключая имя разработчика
    });

    // Get data for all developers
    let allocationData = getAllocationData(developers, projects, isLastWeek);
    let allAllocationData = allocationData.allocationData;
    let allDailyData = allocationData.dailyData;

    developers = getDevelopers(workloadSheet, all, allAllocationData, allDailyData, isLastWeek);

    if(all) {
        // Обходим данные allAllocationData и добавляем недостающие проекты и разработчиков
        for (let developerName in allAllocationData) {
            let developerIndex = developers.findIndex(developer => developer.name === developerName);

            if (developerIndex === -1) {
                // Добавляем нового разработчика
                developers.push({
                    name: developerName,
                    location: '', // Местоположение нам неизвестно, выставляем пустую строку
                    projectHours: {}, // Часы по проектам устанавливаем как пустой объект
                    projects: {}, // Проекты устанавливаем как пустой объект
                    vacationHours: 0, // Часы отпуска устанавливаем в 0
                });
                developerIndex = developers.length - 1; // Update the developer index to the newly added developer
            }

            let developerProjects = allAllocationData[developerName].projects;
            for (let projectName in developerProjects) {
                let projectIndex = projects.findIndex(project => project.projectName === projectName);

                if (projectIndex === -1) {
                    // Добавляем новый проект
                    projects.push({
                        pmInitials: '', // Инициалы менеджера нам неизвестны, выставляем пустую строку
                        projectName: projectName,
                        projectHours: 0, // Часы по проектам устанавливаем в 0
                        developers: {}, // Разработчики устанавливаем как пустой объект
                    });
                    projectIndex = projects.length - 1; // Update the project index to the newly added project
                }

                // Add new developer to the project's developers list if not already present
                if (!projects[projectIndex].developers[developerName]) {
                    projects[projectIndex].developers[developerName] = 0;
                }

                // Add new project to the developer's projects list if not already present
                if (!developers[developerIndex].projects[projectName]) {
                    developers[developerIndex].projects[projectName] = 0;
                }

                // Set hours for the new developer's project to 0 if not already present
                if (!developers[developerIndex].projectHours[projectName]) {
                    developers[developerIndex].projectHours[projectName] = 0;
                }
            }
        }
    }

    // Шаг 1. Получите данные проекта "vacation".
    let vacationData = developers.map(developer => {
        return [developer.name, null, null, "vacation", developer.vacationHours || 0];
    }).filter(row => row[4] > 0); // Фильтруйте разработчиков с нулевыми часами отпуска

    // Шаг 2. Добавьте данные проекта "vacation" в allAllocationData.
    vacationData.forEach(([developerName, , , project, hours]) => {
        // Если нет данных для этого разработчика, создайте их
        if (!allAllocationData[developerName]) {
            allAllocationData[developerName] = {projects: {}, list: ''};
        }

        // Добавьте часы отпуска к проекту "vacation"
        if (!allAllocationData[developerName].projects[project]) {
            allAllocationData[developerName].projects[project] = 0;
        }
        allAllocationData[developerName].projects[project] += hours;
        if (!allAllocationData[developerName].list) {
            hours = hours.toFixed(2);
            allAllocationData[developerName].list += "ВМ vacation (" + hours + ")";
        }
    });


    var getStackData = (function() {
        var allDevelopersStackData = null;  // Закрытая переменная

        return function(developerName) {
            if (!allDevelopersStackData) {
                allDevelopersStackData = getAllDevelopersStackDataFromSheet();  // Получить данные только при первом вызове
            }

            // Возвращаемые данные для указанного разработчика или undefined, если такого разработчика нет
            return allDevelopersStackData[developerName] || {};
        }
    })();

    var getCompetenceData = (function() {
        var developerCompetenceData = null;  // Закрытая переменная

        return function(developerName) {
            if (!developerCompetenceData) {
                developerCompetenceData = getAllDevelopersCompetenceData();  // Получить данные только при первом вызове
            }

            // Возвращаемые данные для указанного разработчика или undefined, если такого разработчика нет
            return developerCompetenceData[developerName] || {};
        }
    })();

    var getDeveloperCvList = (function() {
        var allDevelopersCvData = null;  // Private variable

        return function(developerName) {
            if (!allDevelopersCvData) {
                allDevelopersCvData = getAllDevelopersCvDataFromSheet();  // Get data only on the first call
            }

            // Returned data for the specified developer or an empty object if such a developer does not exist
            return allDevelopersCvData[developerName] || { folders: {} };
        }
    })();

    var getDeveloperUpworkData = (function() {
        var allDevelopersUpworkData = null;  // Закрытая переменная

        return function(russianName) {
            if (!allDevelopersUpworkData) {
                allDevelopersUpworkData = getAllDevelopersUpworkDataFromSheet();  // Получить данные только при первом вызове
            }

            // Возвращаемые данные для указанного разработчика или пустой объект, если такого разработчика нет
            return allDevelopersUpworkData[russianName] || {};
        }
    })();

    var getCandidateData = (function() {
        var allCandidatesData = null;  // Закрытая переменная

        return function(candidateName) {
            if (!allCandidatesData) {
                allCandidatesData = getAllCandidatesDataFromSheet();  // Получить данные только при первом вызове
            }

            // Возвращаемые данные для указанного кандидата или пустой объект, если такого кандидата нет
            return allCandidatesData[candidateName] || {};
        }
    })();



    //// Logger.log(developers.length);


    // Initialize report



    if(!isBench)
        reportSheet = ss.getSheetByName(reportName);
    else
        reportSheet = benchSpreadsheet.getSheetByName(reportName);


    let currentReportName = reportSheet.getRange('B3').getValue();

    // if (!isBench && currentReportName !== (reportName + ` for ${mondayString} - ${sundayString}`)) {
    //     // Если имя отчета отличается, архивируем текущий лист
    //     let archivedSheet = reportSheet.copyTo(ss);

    //     // Переименовываем архивный лист и перемещаем его в конец
    //     archivedSheet.setName(currentReportName);
    //     ss.setActiveSheet(archivedSheet);
    //     ss.moveActiveSheet(ss.getNumSheets());
    // }


    var range = reportSheet.getDataRange();
    var notes = range.getNotes();
    for (var i = 0; i < notes.length; i++) {
        for (var j = 0; j < notes[i].length; j++) {
            var cellNotes = notes[i][j];
            if (cellNotes !== "") {
                reportSheet.getRange(i + 1, j + 1).clearNote();
            }
        }
    }

    // Проверяем, есть ли скрытые строки
    var maxRows = reportSheet.getLastRow();
    var hiddenRows = [];
    for (var i = 1; i <= maxRows; i++) {
        if (reportSheet.isRowHiddenByUser(i)) {
            hiddenRows.push(i);
        }
    }

    // Если есть скрытые строки, показываем их
    if (hiddenRows.length > 0) {
        reportSheet.showRows(1, maxRows);
    }

    // Если есть активный фильтр, удаляем его
    var filter = reportSheet.getFilter();
    if (filter) {
        filter.remove();
    }

    reportSheet.clear();
    reportSheet.clearContents();



    reportSheet.getRange('B3').setValue( reportName + ` for ${mondayString} - ${sundayString}`).setFontSize(20);
    // reportSheet.getRange('K3').setValue('для сортировки выделите колонку и нажмите "Сортировать" или используйте дополнительные инструменты поиска в меню "Фильтры"').setFontSize(9);

    let upworkCVwidth = 150;
    if(isBench) upworkCVwidth = 1;

    let column = 2;

    // availableHours column
    var textColumn = 14;
    if (isBench) textColumn = 8;
    if (isLastWeek) textColumn = 10;

    // Initialize the header row
    reportSheet.getRange(5, column).setValue('Developer').setVerticalAlignment("middle");
    reportSheet.setColumnWidth(column, 200);
    column++;
    reportSheet.getRange(5, column).setValue('Location').setVerticalAlignment("middle");
    reportSheet.setColumnWidth(column, 200);
    column++;
    reportSheet.getRange(5, column).setValue('English').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(column, 30);
    column++;
    reportSheet.getRange(5, column).setValue('Training').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(column, 30);
    column++;
    reportSheet.getRange(5, column).setValue('Sales').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(column, 30);
    column++;


    if(!isBench && !isLastWeek) {
        reportSheet.getRange(5, column).setValue(workloadSheetName).setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
        reportSheet.setColumnWidth(column, 40);
        column++;
        reportSheet.getRange(5, column).setValue(workloadNextWeekSheetName).setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
        reportSheet.setColumnWidth(column, 40);
        column++;
        reportSheet.getRange(5, column).setValue(workloadNextNextWeekSheetName).setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
        reportSheet.setColumnWidth(column, 40);
        column++;
        reportSheet.getRange(5, column).setValue(workloadNextNextNextWeekSheetName).setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
        reportSheet.setColumnWidth(column, 40);
        column++;
    }

    if(!isBench) {
        reportSheet.getRange(5, column).setValue('CV').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
        reportSheet.setColumnWidth(column, upworkCVwidth);
        column++;
        reportSheet.getRange(5, column).setValue('Upwork').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
        reportSheet.setColumnWidth(column, upworkCVwidth);
        column++;
    }


    reportSheet.getRange(5, column).setValue('Stack').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(column, 150);
    column++;
    reportSheet.getRange(5, column).setValue('Learning ability\nManaging stress\nTeam work\nClient communication\nPresentation skills\nThinking speed').setTextRotation(90).setBackground("#ffffff").setHorizontalAlignment("center").setVerticalAlignment("middle");
    reportSheet.setColumnWidth(column, 150);
    column++;

    let allStacks = {};

    for (let developer of developers) {
        if(!developer.name) continue;
        let developerName = developer.name.split("(")[0].trim();
        let stackData = getStackData(developerName);

        for (let stack in stackData) {
            if(stack != '') {
                //stack = stack.toUpperCase();
                if (!allStacks.hasOwnProperty(stack)) {
                    allStacks[stack] = 0;
                }
                allStacks[stack]++;
            }
        }
    }
    let sortedStacks = Object.keys(allStacks).sort((a, b) => allStacks[b] - allStacks[a]);

    let n = 0;
    for (let stack of sortedStacks) {
        //stack = stack.toUpperCase();
        n++;
        reportSheet.getRange(5, column).setValue(stack).setVerticalAlignment("middle").setHorizontalAlignment("center").setTextRotation(90).setBackground("#cccccc").setFontSize(9);
        reportSheet.setColumnWidth(column, 25);
        column++;
    }

    let columnForTable = column;

    if(!isBench) {
        reportSheet.getRange(5,column).setValue('Plan').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
        reportSheet.setColumnWidth(column, 150);
        column++;
        reportSheet.getRange(5,column).setValue('Fact').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
        reportSheet.setColumnWidth(column, 150);
        column++;


        // Write TOTAL in the next two columns
        reportSheet.getRange(5, column)
            .setValue('TOTAL plan')
            .setVerticalAlignment("middle")
            .setHorizontalAlignment("center")
            .setTextRotation(90)
            .setBackground("#ffffff")
            .setFontSize(9);
        column++;
        reportSheet.setColumnWidth(column, 40);

        // Leave a column for 'fact' data
        reportSheet.getRange(5, column)
            .setValue('TOTAL fact')
            .setVerticalAlignment("middle")
            .setHorizontalAlignment("center")
            .setTextRotation(90)
            .setBackground("#ffffff")
            .setFontSize(9);
        column++;
        // Add a border to the right of the empty column

        reportSheet.getRange(5, column, 120).setBorder(false, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        reportSheet.getRange(3, column).setValue(sumColumnE);
        column++;

    }


    // if(all) {


    //     for (let project of projects) {
    //         // Write the project name and PM initials in the next two columns
    //         reportSheet.getRange(5, column)
    //             .setValue(project.pmInitials + ' ' + project.projectName + ' plan')
    //             .setVerticalAlignment("middle")
    //             .setHorizontalAlignment("center")
    //             .setTextRotation(90)
    //             .setBackground("#cccccc")
    //             .setFontSize(9);
    //         reportSheet.setColumnWidth(column, 40);

    //         column++;

    //         // Leave a column for 'fact' data
    //         reportSheet.getRange(5, column)
    //             .setValue(project.pmInitials + ' ' + project.projectName + ' fact')
    //             .setVerticalAlignment("middle")
    //             .setHorizontalAlignment("center")
    //             .setTextRotation(90)
    //             .setBackground("#cccccc")
    //             .setFontSize(9);
    //         reportSheet.setColumnWidth(column, 40);

    //         column++;

    //         // Skip an empty column
    //         reportSheet.getRange(5, column).setBackground("#ffffff");
    //         reportSheet.setColumnWidth(column, 40);

    //         // Add a border to the right of the empty column
    //         reportSheet.getRange(5, column, 120).setBorder(false, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    //         // Increment the column counter to skip the 'fact' column
    //         column++;
    //     }

    // }


    if (!isBench) {
        reportSheet.getRange(5, column).setValue('Available vacation').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
        column++;
        reportSheet.getRange(5, column).setValue('Hourly Rate').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
        column++;
        reportSheet.getRange(5, column).setValue('Projects').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
        column++;
        reportSheet.getRange(5, column).setValue('Create CV from Profile').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    }

    // Initialize the data rows
    let row = 7;
    for (let developer of developers) {
        if(!developer.name) continue;
        let developerName = developer.name.split("(")[0].trim(); // Remove everything after the "(" and trim spaces

        Logger.log('Output a row for ' + developerName);

        let allocationList = '';
        if(allAllocationData[developerName] && allAllocationData[developerName].list) allocationList = allAllocationData[developerName].list;

        let developerAllocationData = allAllocationData[developerName];
        // if (developerAllocationData && developerAllocationData.list) {
        //   allocationList = developerAllocationData.list;
        // }

        //let competenceData = developerCompetenceData(developerName)
        let englishLevel = getCompetenceData(developerName)['Английский'];
        // Here you need to calculate trainingAndSales and allocation for each developer
        let trainingHours = developer.projects['Training'] ? developer.projects['Training'] : 0;
        let salesHours = developer.projects['SALES'] ? developer.projects['SALES'] : 0;
        let weekPlanHoursTotal = 0;
        let weekProjectHours = '';
        let nextWeekPlanHoursTotal = 0;
        let nextWeekProjectHours = '';
        let nextNextWeekPlanHoursTotal = 0;
        let nextNextWeekProjectHours = '';
        let nextNextNextWeekPlanHoursTotal = 0;
        let nextNextNextWeekProjectHours = '';

        let stackData = getStackData(developerName);

        if(!isBench) {
            var developerCvData = getDeveloperCvList(developerName);
            var combinedRichTextBuilder = SpreadsheetApp.newRichTextValue();
            var combinedText = '';

            var developerUpworkData = getDeveloperUpworkData(developerName);
            var candidateData = getCandidateData(developerName);

            // Iterate over each folder
            for (var folderId in developerCvData.folders) {
                var developerCvList = developerCvData.folders[folderId].cvList;

                // Add a link to the folder
                var cvFolderLink = 'https://drive.google.com/drive/folders/' + folderId;
                combinedText += 'CV folder: ' + folderId + '\n';

                // Add text for each CV in the folder
                developerCvList.forEach(function(cv) {
                    var date = Utilities.formatDate(new Date(cv.lastUpdate), 'GMT', 'dd/MM/yyyy');  // Convert the date to the format dd/MM/yyyy
                    var linkText = cv.fileName + '\n';  // Link text
                    var text = date + '\n';  // Update date
                    combinedText += linkText + text;
                });
            }

            combinedRichTextBuilder.setText(combinedText);  // Update the text in RichTextValueBuilder

            // Add links for each CV
            var index = 0;
            for (var folderId in developerCvData.folders) {
                var developerCvList = developerCvData.folders[folderId].cvList;

                // Add a link to the folder
                var cvFolderLink = 'https://drive.google.com/drive/folders/' + folderId;
                combinedRichTextBuilder.setLinkUrl(index, index + ('CV folder: ' + folderId).length, cvFolderLink);  // Link to the CV folder
                index += ('CV folder: ' + folderId + '\n').length;

                // Add links for each CV in the folder
                developerCvList.forEach(function(cv) {
                    var linkText = cv.fileName + '\n';  // Link text
                    var text = Utilities.formatDate(new Date(cv.lastUpdate), 'GMT', 'dd/MM/yyyy') + '\n';  // Text with the date
                    var fullText = linkText + text;  // Full text
                    combinedRichTextBuilder.setLinkUrl(index, index + linkText.length - 1, cv.link);  // Link to the CV
                    index += fullText.length;
                });
            }

            // Build the final RichTextValue
            var cvDataRichText = combinedRichTextBuilder.build();


            // Создание RichTextValueBuilder
            var combinedRichTextBuilder2 = SpreadsheetApp.newRichTextValue();
            var combinedText = '';

            // Добавление данных Upwork
            var upworkText = '';
            for (var key in developerUpworkData) {
                if (developerUpworkData.hasOwnProperty(key)) {
                    var value = developerUpworkData[key];
                    if (value) { // Проверяем, что значение не пустое
                        upworkText += key.charAt(0).toUpperCase() + key.slice(1) + ': ' + value + '\n';
                    }
                }
            }
            combinedText += upworkText; // Добавляем данные Upwork в общий текст

            combinedRichTextBuilder2.setText(combinedText); // Обновляем текст в RichTextValueBuilder

            // Добавляем ссылку на профиль Upwork, если она есть
            if (developerUpworkData.upworkLink) {
                var linkStart = combinedText.indexOf(developerUpworkData.upworkLink);
                var linkEnd = linkStart + developerUpworkData.upworkLink.length;
                combinedRichTextBuilder2.setLinkUrl(linkStart, linkEnd, developerUpworkData.upworkLink);
            }

            // Построение итогового RichTextValue
            var upworkRichText = combinedRichTextBuilder2.build();

            var candidateComment = '';

            var fieldsToInclude = ["Кандидат полностью обработан\nДата (заполнить)", "Вакансия", "HR Commentary", "Tech Stack", "Skype/Telegram", "Interview comments"];

            for (var i = 0; i < candidateData.length; i++) {
                if (candidateData[i].value && fieldsToInclude.includes(candidateData[i].field)) { // Если значение не пустое и поле нужно включить
                    candidateComment += candidateData[i].field + ':\n' + candidateData[i].value + '\n\n';
                }
            }
        }

        if(!isBench && !isLastWeek) {
            for (let week in nextWeekProjects) {
                for (let project of nextWeekProjects[week]) {
                    for (let developer in project.developers) {
                        if (developer.startsWith(developerName) && project.developers[developer]>0) {
                            if(week === 'week') {
                                weekPlanHoursTotal += project.developers[developer];
                                weekProjectHours += project.pmInitials + ' ' + project.projectName + ' (' + project.developers[developer] + ')\n';
                            } else if(week === 'nextWeek') {
                                nextWeekPlanHoursTotal += project.developers[developer];
                                nextWeekProjectHours += project.pmInitials + ' ' + project.projectName + ' (' + project.developers[developer] + ')\n';
                            } else if(week === 'nextNextWeek') {
                                nextNextWeekPlanHoursTotal += project.developers[developer];
                                nextNextWeekProjectHours += project.pmInitials + ' ' + project.projectName + ' (' + project.developers[developer] + ')\n';
                            } else if(week === 'nextNextNextWeek') {
                                nextNextNextWeekPlanHoursTotal += project.developers[developer];
                                nextNextNextWeekProjectHours += project.pmInitials + ' ' + project.projectName + ' (' + project.developers[developer] + ')\n';
                            }
                        }
                    }
                }
            }

            if (weekPlanHoursTotal > 0 && trainingHours > 0) weekPlanHoursTotal = weekPlanHoursTotal - trainingHours;
        }

        // для SALES нужны только они
        if ((!all || isBench) && nextWeekPlanHoursTotal >= 20 && nextNextWeekPlanHoursTotal >= 20 && nextNextNextWeekPlanHoursTotal >= 20 && trainingHours === 0 && salesHours === 0) {
            Logger.log('Пропускаем ' + developerName);
            continue;
        }

        let column = 2;
        let developerNameToShow = isBench ? transliterate(developerName) : developerName;
        let englishName = findEnglishName(developerNames, developerName);
        let profileLink = findDeveloperProfileLink(developerProfiles, englishName);
        let profileId = profileLink !== -1 ? extractIdFromUrl(profileLink) : '';

        if (profileLink !== -1 && profileLink !== '' && !isBench) {
            reportSheet.getRange(row, 2).setFormula(`=HYPERLINK("${profileLink}", "${developerNameToShow}")`).setVerticalAlignment("middle");
        } else if (all || isLastWeek) {
            reportSheet.getRange(row, 2).setValue(developerNameToShow).setVerticalAlignment("middle").setWrap(true);
        } else {
            continue;
        }

        reportSheet.getRange(row, column).setNote(candidateComment);
        column++;
        reportSheet.getRange(row, column).setValue(developer.location).setVerticalAlignment("middle").setWrap(true);
        column++;
        reportSheet.getRange(row, column).setValue(englishLevel).setVerticalAlignment("middle");
        column++;
        reportSheet.getRange(row, column).setValue(trainingHours).setVerticalAlignment("middle");
        column++;
        reportSheet.getRange(row, column).setValue(salesHours).setVerticalAlignment("middle");
        column++;

        if(!isBench && !isLastWeek) {
            reportSheet.getRange(row, column).setValue(weekPlanHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center").setNote(weekProjectHours);
            if (weekPlanHoursTotal <= 20) reportSheet.getRange(row, column).setBackground("#d9ead3");
            else reportSheet.getRange(row, column).setBackground("#ffffff");
            column++;
            reportSheet.getRange(row, column).setValue(nextWeekPlanHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center").setNote(nextWeekProjectHours);
            if (nextWeekPlanHoursTotal <= 20) reportSheet.getRange(row, column).setBackground("#d9ead3");
            else reportSheet.getRange(row, column).setBackground("#ffffff");
            column++;
            reportSheet.getRange(row, column).setValue(nextNextWeekPlanHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center").setNote(nextNextWeekProjectHours);
            if (nextNextWeekPlanHoursTotal <= 20) reportSheet.getRange(row, column).setBackground("#d9ead3");
            else reportSheet.getRange(row, column).setBackground("#ffffff");
            column++;
            reportSheet.getRange(row, column).setValue(nextNextNextWeekPlanHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center").setNote(nextNextNextWeekProjectHours);
            if (nextNextNextWeekPlanHoursTotal <= 20) reportSheet.getRange(row, column).setBackground("#d9ead3");
            else reportSheet.getRange(row, column).setBackground("#ffffff");
            column++;

            if (trainingHours >= 10 || weekPlanHoursTotal <= 20) {
                // Выделить строку зеленым цветом
                reportSheet.getRange(row, 2, 1, textColumn - 5).setBackground("#d9ead3"); // Смените число 11 на число столбцов в вашей строке
            }
        }



        if(!isBench) {
            if (cvDataRichText) {
                // Присваивание RichTextValue ячейке
                reportSheet.getRange(row, column).setRichTextValue(cvDataRichText).setVerticalAlignment("middle").setHorizontalAlignment("left").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setFontSize(8);
            }
            column++;

            if (developerUpworkData.upworkLink) {
                // Присваивание обычного текста ячейке
                reportSheet.getRange(row, column).setRichTextValue(upworkRichText).setVerticalAlignment("middle").setHorizontalAlignment("left").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setFontSize(8);
                reportSheet.getRange(row, column).setNote(upworkText);
            }
            column++;
        }

        let stackText = 'Main\n' + (getCompetenceData(developerName)['Основной стек'] ?? '') + '\n\nExtra\n' + (getCompetenceData(developerName)['Дополнительный стек'] ?? '');
        let competenceText = getCompetenceData(developerName)['Инструменты\nБиблиотеки\nСитстемы'] ?? '';
        reportSheet.getRange(row, column).setValue(stackText).setNote(stackText + '\n\n' + 'Skills\n' + competenceText).setVerticalAlignment("middle").setFontSize(8);
        column++;

        reportSheet.getRange(row, column).setValue(
            (getCompetenceData(developerName)['Обучаемость'] ?? '') + '  ' +
            (getCompetenceData(developerName)['Стрессоустойчивость'] ?? '') + '  ' +
            (getCompetenceData(developerName)['Работа в команде'] ?? '') + '  ' +
            (getCompetenceData(developerName)['Работа с клиентом (командой клиента)'] ?? '') + '  ' +
            (getCompetenceData(developerName)['Навыки самопрезентации'] ?? '') + '  ' +
            (getCompetenceData(developerName)['Гибкость мышления'] ?? '')
        ).setVerticalAlignment("middle").setHorizontalAlignment("center");


        //column++;

        // n = 0;

        Logger.log('Выводим стеки для ' + developerName);

        // for (let stack of sortedStacks) {
        //     n++;
        //     let stackLevel = stackData[stack] || '';
        //     let cell = reportSheet.getRange(row, column);
        //     cell.setValue(stackLevel).setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");

        //     // Устанавливаем цвета для разных уровней
        //     if (stackLevel.startsWith('jun')) {
        //         cell.setBackground("#add8e6");  // Цвет для Junior (Светло-синий)
        //     } else if (stackLevel.startsWith('mid')) {
        //         cell.setBackground("#90ee90");  // Цвет для Middle (Светло-зелёный)
        //     } else if (stackLevel.startsWith('sr')) {
        //         cell.setBackground("#f4a460");  // Цвет для Senior (Светло-коричневый)
        //     }
        //     column++;
        // }


        /// optimized stacks

        let values = [];
        let styles = [];
        n = 0;  // Переместим объявление сюда

        for (let stack of sortedStacks) {
            n++;
            let stackLevel = stackData[stack] || '';
            values.push(stackLevel);  // Изменяем на одномерный массив

            let style = {
                textRotation: 90,
                verticalAlignment: "middle",
                horizontalAlignment: "center",
                background: ""
            };

            if (stackLevel.startsWith('jun')) {
                style.background = "#add8e6";  // Цвет для Junior (Светло-синий)
            } else if (stackLevel.startsWith('mid')) {
                style.background = "#90ee90";  // Цвет для Middle (Светло-зелёный)
            } else if (stackLevel.startsWith('sr')) {
                style.background = "#f4a460";  // Цвет для Senior (Светло-коричневый)
            }

            styles.push(style);
            column++;
        }

        let startColumn = column - sortedStacks.length + 1;
        let range = reportSheet.getRange(row, startColumn, 1, sortedStacks.length);

        // Преобразование values в двумерный массив для setValues
        let valuesToSet = [values];
        range.setValues(valuesToSet);

        range.setTextRotation(90)
            .setVerticalAlignment("middle")
            .setHorizontalAlignment("center");

        for (let i = 0; i < sortedStacks.length; i++) {
            range.getCell(1, i + 1).setBackground(styles[i].background);
        }

        column++;

        ///////////////////////////////////////////////


        if (!isBench) {
            reportSheet.getRange(row, column).setValue(developer.projectHours).setVerticalAlignment("top").setWrap(true).setFontSize(8);
            reportSheet.getRange(row, column+1).setValue(allocationList).setVerticalAlignment("top").setWrap(true).setFontSize(8);

            let columnPlanHours = column+2;
            let columnFactHours = column+3;
            let columnDiffHours = column+4;

            column = column+5;


            let planHoursTotal = 0;
            let factHoursTotal = 0;

            for (let project of projects) {
                let dataRow = [];

                // Write the plan hours for the developer for this project
                let planHours = Object.keys(project.developers).find(devName => devName.startsWith(developer.name));
                planHours = project.developers[planHours] || '';


                // Write the fact hours for the developer for this project
                let factHours = '';
                if (developerAllocationData && developerAllocationData.projects) {
                    factHours = developerAllocationData.projects[project.projectName] || '';
                }

                // if(all) {
                //     dataRow.push({
                //         value: planHours,
                //         verticalAlignment: "middle",
                //         horizontalAlignment: "center",
                //         background: "#cccccc",
                //         fontSize: 8
                //     });

                //     dataRow.push({
                //         value: factHours,
                //         verticalAlignment: "middle",
                //         horizontalAlignment: "center",
                //         background: "#cccccc",
                //         fontSize: 8
                //     });

                //     // Calculate the difference (plan - fact) and write in the next column
                //     let formula = `=IF(AND(ISBLANK(R${row}C${column}), ISBLANK(R${row}C${column+1})), "", R${row}C${column+1}-R${row}C${column})`;
                //     let difference = (factHours - planHours) || '';
                //     let color = difference < 0 ? "red" : "green";
                //     dataRow.push({
                //         value: difference,
                //         formula: formula,
                //         verticalAlignment: "middle",
                //         horizontalAlignment: "center",
                //         background: "#ffffff",
                //         fontSize: 8,
                //         fontColor: color
                //     });

                //     // Write all the data to the row at once
                //     let range = reportSheet.getRange(row, column, 1, 3);
                //     range.setValues([dataRow.map(cell => cell.value)]);
                //     range.setBackgrounds([dataRow.map(cell => cell.background)]);
                //     range.setFontColors([dataRow.map(cell => cell.fontColor)]);
                //     range.setFontSizes([dataRow.map(cell => cell.fontSize)]);
                //     range.setVerticalAlignments([dataRow.map(cell => cell.verticalAlignment)]);
                //     range.setHorizontalAlignments([dataRow.map(cell => cell.horizontalAlignment)]);

                //     // Skip an empty column
                //     reportSheet.getRange(5, column + 2).setBackground("#ffffff");
                //     reportSheet.setColumnWidth(column + 2, 35);

                //     // Increment the column counter to skip the 'fact' column
                //     column += 3;
                // }



                planHours = Math.round(planHours * 100) / 100;
                factHours = Math.round(factHours * 100) / 100;

                planHoursTotal += planHours;
                factHoursTotal += factHours;


                let diffHoursTotal = factHoursTotal-planHoursTotal;
                let diffFontColor = "green"
                if(diffHoursTotal < 0) diffFontColor = "red";

                reportSheet.getRange(row, columnPlanHours).setValue(planHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center").setNote(developer.projectHours);
                reportSheet.getRange(row, columnFactHours).setValue(factHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center").setNote(allocationList);
                reportSheet.getRange(row, columnDiffHours).setValue(diffHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center").setFontColor(diffFontColor);

            }

            Logger.log('Выводим оставшиеся выходные дни и ставки ' + developerName);

            let vacationDays = findDeveloperVacation(developerVacations, developerName);
            if (vacationDays !== -1) {
                var days = Math.floor(vacationDays[1]);
                if (days < 0 || days > 0) {
                    reportSheet.getRange(row, column).setValue(days).setVerticalAlignment("middle").setHorizontalAlignment("center");
                }
            }
            column++;

            let hourlyRates = findDeveloperRate(developerRates, developerName);
            if (hourlyRates) {
                let hourlyRate = Math.floor(hourlyRates[1]);
                if (hourlyRate > 0) {
                    reportSheet.getRange(row, column).setValue(hourlyRate).setVerticalAlignment("middle").setHorizontalAlignment("center");
                }

            }
            column++;

            Logger.log('Выводим информацию по проектам и ссылку на генерацию CV ' + developerName);

            if (profileId !== -1 && profileId !== '') {
                // Добавление информации о проектах в комментарий
                if (developersProjects[englishName]) {
                    var devProjects = developersProjects[englishName];
                    var projectsNumber = devProjects.length;
                    var projectsList = devProjects.map(project => project.join(" | ")).join("\n\n======================\n\n");
                    var projectsCell = reportSheet.getRange(row, column);

                    projectsCell.setValue(projectsNumber);
                    projectsCell.setHorizontalAlignment("center"); // Горизонтальное выравнивание текста по центру
                    projectsCell.setVerticalAlignment("middle"); // Вертикальное выравнивание текста по центру
                    projectsCell.setNote(`${projectsList}`);
                }

                column++;

                var url = 'https://script.google.com/a/macros/sharp-dev.net/s/AKfycbyQRlX26I41ekAF8uc2qY3VrqEui2tLdNcx81gwP_wY44IrWz-D2O_Nndajqvbf-5ZN/exec?documentId=' + profileId;
                var formula = '=HYPERLINK("' + url + '", "Generate CV")';
                var cell = reportSheet.getRange(row, column);
                cell.setFormula(formula);

                // Форматирование ячейки для имитации кнопки
                cell.setBackground("#f4f4f4"); // Светло-серый фон
                cell.setFontColor("#1a73e8"); // Цвет текста, как у стандартных ссылок Google
                cell.setFontWeight("bold"); // Жирный шрифт
                cell.setHorizontalAlignment("center"); // Горизонтальное выравнивание текста по центру
                cell.setVerticalAlignment("middle"); // Вертикальное выравнивание текста по центру
                cell.setBorder(true, true, true, true, false, false, "#cccccc", SpreadsheetApp.BorderStyle.SOLID_MEDIUM); // Рамка вокруг ячейки

                // Установка высоты и ширины ячейки для большего сходства с кнопкой (опционально, зависит от вашего макета)
                // reportSheet.setRowHeight(row, 35); // Установка высоты строки
                reportSheet.setColumnWidth(column, 100); // Установка ширины столбца


            }

            row++;
        } else {
            row++
        }





    }


    // Set the border
    reportSheet.getRange(5, 2, row-5, columnForTable-2).setBorder(true, true, true, true, true, true);

    insertSumFormulas(all,isLastWeek, isBench);
    let lastColumn = reportSheet.getLastColumn();

    // определите номер строки, куда нужно вставить итоговые значения (после последней строки с данными)
    let totalRow = reportSheet.getLastRow() + 1;


    reportSheet.getRange(4, textColumn).setValue('available hours:');

    // Get the values in the fifth row
    let fifthRowValues = reportSheet.getRange(5, 1, 1, reportSheet.getLastColumn()).getValues()[0];

    // Initialize endColumn with the length of the fifthRowValues
    let endColumn = fifthRowValues.length;

    // Look for the first set of three empty cells in the fifth row
    for (let i = 0; i < fifthRowValues.length - 2; i++) {
        if (fifthRowValues[i] === 'Plan') {
            endColumn = i;
            break;
        }
    }

    // так как у бенча нет колонок с планом и фактом.
    if(isBench) endColumn = lastColumn;

    // Apply the first formula from startColumn to endColumn
    for(let i = textColumn + 1; i <= endColumn; i++) {
        let columnLetter = getColumnLetter(i);
        let formula = `=SUMIF(XXX7:XXX${totalRow-1}, "<>", A7:A${totalRow-1})`;
        formula = formula.replace('XXX7:XXX', `${columnLetter}7:${columnLetter}`);
        formula = formula.replace('XXX7:A', `${columnLetter}7:A`);
        reportSheet.getRange(4, i).setFormula(formula);
    }


    if(all) {
        // The starting column for your formula for the remaining columns after the 3 empty ones
        let remainingStartColumn = endColumn + 3;  // Adding 3 to account for fact and plan columns and the column where the next set of data begins

        // Apply the second formula from remainingStartColumn to the last column
        let counter = 0;
        for(let i = remainingStartColumn; i <= lastColumn; i++) {
            let columnLetter = getColumnLetter(i);
            let formula = `=SUM(K7:K${totalRow-1})`;
            formula = formula.replace('K7:K', `${columnLetter}7:${columnLetter}`);

            let cell = reportSheet.getRange(4, i);
            cell.setFormula(formula).setFontSize(8).setHorizontalAlignment("center");

            // Make sure the formula is evaluated
            SpreadsheetApp.flush();

            counter++;

            // If it's the third cell, get its value and set the color accordingly
            if (counter % 3 == 0) {
                let cellValue = cell.getValue();
                if (cellValue < 0) {
                    cell.setFontColor("red");
                } else {
                    cell.setFontColor("green");
                }
            }
        }
    }

    // Если есть активный фильтр, удаляем его
    var filter = reportSheet.getFilter();
    if (filter) {
        filter.remove();
    }

    // создаем новый фильтр
    if ((sheet.getLastRow() - 5) > 6) var range = reportSheet.getRange(6, 2, sheet.getLastRow() - 5, lastColumn-1);
    if (range) range.createFilter();

    // Add date and time of data gathering
    const currentTime = new Date().toLocaleString("en-GB", {timeZone: "Asia/Tbilisi"});
    reportSheet.getRange("B4").setValue(`Generated at ${currentTime} (Tbilisi, Georgia Timezone)`);

    // Определите диапазон строк, которому вы хотите задать новую высоту
    reportSheet.setRowHeights(1, 1, 1);
    reportSheet.setRowHeights(2, 1, 1);
    reportSheet.setRowHeights(3, 1, 1);
    reportSheet.setRowHeights(4, 1, 20);
    reportSheet.setRowHeights(5, 1, 80);
    reportSheet.setRowHeights(6, 1, 20);

}


function generateWeekReportCurrentWeek() {
    generateWeekReport(false);
}

function generateWeekReportLastWeek() {
    generateWeekReport(true);
}

function generateWeekReportLast2Weeks() {
    generateWeekReport(true, 2);
}

function generateWeekReportLast2WeeksAllHours() {
    generateWeekReport(true, 2, false);
}

function generateWeekReportLast4Weeks() {
    generateWeekReport(true, 4);
}

function generateWeekReportLast8Weeks() {
    generateWeekReport(true, 8);
}

function generateWeekReportLast12Weeks() {
    generateWeekReport(true, 12);
}

function generateWeekReportLast26Weeks() {
    generateWeekReport(true, 26);
}

function generateWeekReportLast52Weeks() {
    generateWeekReport(true, 52);
}

function generateWeekReport(isLastWeek = false, weeks = 1, isPaidHours = true) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let reportName = "PAID HOURS report";

    const darkRed = "#cb4335"; // Более темный красный цвет
    const lightRed = "#f4c7c3"; // Светлый красный цвет
    const lightYellow = "#ffffe0"; // Светло-желтый цвет

    // Если флаг isLastWeek установлен в true, добавляем 'last week' к имени отчета
    if(isLastWeek) {
        if (weeks === 1) reportName += ' last week';
        else reportName += ` last ${weeks} weeks`;
    }

    // // Проверяем, запущена ли функция на правильном листе
    // if (reportName !== 'PAID HOURS report' && reportName !== 'PAID HOURS report last week') {
    //     // Logger.log('Эта функция может быть запущена только на листe "PAID HOURS report".');
    //     return;
    // }

    const workloadSheetId = "1N65NUtqBA855C6K8swmeFQ9HbvIZU4fq4EnhYzvNV7Q";
    const workloadSpreadsheet = SpreadsheetApp.openById(workloadSheetId);

    let mondayDate, sundayDate;

    if (isLastWeek) {
        if (weeks > 1) {
            mondayDate = getLastMonday(weeks);
            sundayDate = getLastSunday();
        } else {
            mondayDate = getLastMonday();
            sundayDate = getLastSunday();
        }
    } else {
        mondayDate = getCurrentMonday();
        sundayDate = getCurrentSunday();
    }

    let mondayString = Utilities.formatDate(mondayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();
    let sundayString = Utilities.formatDate(sundayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();
    //const dayString = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();

    let workloadSheetName = mondayDate.getMonth() === sundayDate.getMonth() ?
        `${mondayString.split(" ")[0]}-${sundayString.split(" ")[0]} ${sundayString.split(" ")[1]}` :
        `${mondayString}-${sundayString}`;

    let workloadSheet = workloadSpreadsheet.getSheetByName(workloadSheetName);
    if (!workloadSheet && weeks === 1) {
        SpreadsheetApp.getUi().alert(`Cannot find sheet "${workloadSheetName}" in the workload spreadsheet.`);
        return;
    } else if (!workloadSheet && weeks > 1) {
        workloadSheet = null;
    }

    let scrumSheetName = '';
    if (isLastWeek) {
        if (weeks > 1) {
            scrumSheetName = `Scrum files 2024`;
        } else {
            scrumSheetName = `Scrum files for last week`;
        }
    } else {
        scrumSheetName = `Scrum files for current week`;
    }

    const scrumSheet = ss.getSheetByName(scrumSheetName);

    // Get data from column E
    let columnEData = scrumSheet.getRange("E:E").getValues();

    // Sum all the values in column E
    let sumColumnE = columnEData.reduce((sum, value) => {
        // Ensure that the value is a number before adding it to the sum
        if (typeof value[0] == "number") {
            return sum + value[0];
        }
        return sum;
    }, 0);


    /////////////////
    // СБОР ДАННЫХ //
    /////////////////

    console.log('СБОР ДАННЫХ');

    let developers = getDevelopers(workloadSheet, true, null, null, isLastWeek, isPaidHours, -weeks);
    let projectsData = getProjects(workloadSheet, null, isLastWeek, isPaidHours, -weeks);
    let projects = projectsData.projects;
    let paidHoursPlan = projectsData.additionalData[0];
    let weekPlans = {};

    if (weeks > 1) {
        let allDevelopers = {};
        let allProjects = {};
        for (let i = 0; i < weeks; i++) {
            let weekNumber = weeks - i;
            let mondayDate = getLastMonday(weekNumber);
            let sundayDate = getLastSunday(weekNumber);
            let mondayString = Utilities.formatDate(mondayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();
            let sundayString = Utilities.formatDate(sundayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();
            let workloadSheetName = mondayDate.getMonth() === sundayDate.getMonth() ?
                `${mondayString.split(" ")[0]}-${sundayString.split(" ")[0]} ${sundayString.split(" ")[1]}` :
                `${mondayString}-${sundayString}`;

            let workloadSheet = workloadSpreadsheet.getSheetByName(workloadSheetName);
            if (!workloadSheet) {
                SpreadsheetApp.getUi().alert(`Cannot find sheet "${workloadSheetName}" in the workload spreadsheet.`);
                return;
            }

            let projectsData = getProjects(workloadSheet, null, isLastWeek, isPaidHours, -weekNumber);
            let projects = projectsData.projects;
            let developers = getDevelopers(workloadSheet, true,  null, null, isLastWeek, isPaidHours, -weekNumber);

            // Инициализация additionalData и извлечение данных
            let additionalData = projectsData.additionalData || [];
            let paidHoursPlan = 0;
            let workingHours = 0;
            if (additionalData.length > 0) {
                paidHoursPlan = additionalData[0] || 0;
                workingHours = additionalData.find(data => data.workingHours).workingHours;
            }

            //let actualWeekNumber = getWeekNumberForDate();

            weekPlans[weekNumber] = { developers, projects, paidHoursPlan, workingHours };

            for (let project of projects) {
                if (!allProjects[project.projectName]) {
                    allProjects[project.projectName] = project;
                } else {
                    allProjects[project.projectName] = mergeProjects(allProjects[project.projectName], project);
                }
            }

            for (let developer of developers) {
                if (!allDevelopers[developer.name]) {
                    allDevelopers[developer.name] = developer;
                } else {
                    allDevelopers[developer.name] = mergeDevelopers(allDevelopers[developer.name], developer);
                }
            }
        }
        projects = Object.values(allProjects);
        developers = Object.values(allDevelopers);
    }
    else {
        developers = getDevelopers(workloadSheet, true, null, null, isLastWeek, isPaidHours);
        let projectsData = getProjects(workloadSheet, null, isLastWeek, isPaidHours);
        projects = projectsData.projects;
        paidHoursPlan = projectsData.additionalData[0];
        workingHours = projectsData.additionalData.find(data => data.workingHours).workingHours;
    }

    let planDevelopers = developers;

    // Get data for all developers
    // if (weeks > 1) workloadSheet = null;

    // getAllocationData(developers, projects, isLastWeek,true, weeks);

    let allocationData = getAllocationData(developers, projects, isLastWeek, isPaidHours, weeks);
    let allAllocationData = allocationData.allocationData;
    let allDailyData = allocationData.dailyData;
    let allScrumData = allocationData.scrumData;

    let pmList = [];

    allScrumData.forEach(row => {
        const [developerShort, date, type, project, hours] = row;

        if (type === 'PM' || type === 'PMfree') {
            if (!pmList.includes(developerShort)) {
                pmList.push(developerShort);
            }
        }
    });

    console.log(pmList);

    developers = getDevelopers(workloadSheet, true, allAllocationData, allDailyData, isLastWeek, isPaidHours);

    // Обходим данные allAllocationData и добавляем недостающие проекты и разработчиков
    for (let developerName in allAllocationData) {
        let developerIndex = developers.findIndex(developer => developer.name === developerName);

        if (developerIndex === -1) {
            // Добавляем нового разработчика
            developers.push({
                name: developerName,
                location: '', // Местоположение нам неизвестно, выставляем пустую строку
                projectHours: {}, // Часы по проектам устанавливаем как пустой объект
                projects: {}, // Проекты устанавливаем как пустой объект
                vacationHours: 0, // Часы отпуска устанавливаем в 0
            });
            developerIndex = developers.length - 1; // Update the developer index to the newly added developer
        }

        let developerProjects = allAllocationData[developerName].projects;
        for (let projectName in developerProjects) {
            let projectIndex = projects.findIndex(project => project.projectName === projectName);

            if (projectIndex === -1) {
                // Добавляем новый проект
                projects.push({
                    pmInitials: '', // Инициалы менеджера нам неизвестны, выставляем пустую строку
                    projectName: projectName,
                    projectHours: 0, // Часы по проектам устанавливаем в 0
                    developers: {}, // Разработчики устанавливаем как пустой объект
                });
                projectIndex = projects.length - 1; // Update the project index to the newly added project
            }

            // Add new developer to the project's developers list if not already present
            if (!projects[projectIndex].developers[developerName]) {
                projects[projectIndex].developers[developerName] = 0;
            }

            // Add new project to the developer's projects list if not already present
            if (!developers[developerIndex].projects[projectName]) {
                developers[developerIndex].projects[projectName] = 0;
            }

            // Set hours for the new developer's project to 0 if not already present
            if (!developers[developerIndex].projectHours[projectName]) {
                developers[developerIndex].projectHours[projectName] = 0;
            }
        }
    }

    // Шаг 1. Получите данные проекта "vacation".
    let vacationData = developers.map(developer => {
        return [developer.name, null, null, "vacation", developer.vacationHours || 0];
    }).filter(row => row[4] > 0); // Фильтруйте разработчиков с нулевыми часами отпуска

    // Шаг 2. Добавьте данные проекта "vacation" в allAllocationData.
    vacationData.forEach(([developerName, , , project, hours]) => {
        // Если нет данных для этого разработчика, создайте их
        // надо убедиться что allAllocationData[developerName] существует иначе будет ошибка
        if (!allAllocationData[developerName]) {
            allAllocationData[developerName] = {projects: {}, list: ''};
        }

        // Добавьте часы отпуска к проекту "vacation"
        if (!allAllocationData[developerName].projects[project]) {
            allAllocationData[developerName].projects[project] = 0;
        }
        allAllocationData[developerName].projects[project] += hours;
        if (!allAllocationData[developerName].list) {
            hours = hours.toFixed(2);
            allAllocationData[developerName].list += "ВМ vacation (" + hours + ")";
        }
    });


    reportSheet = ss.getSheetByName(reportName);
    if (!reportSheet) {
        reportSheet = ss.insertSheet(reportName);
    }

    var range = reportSheet.getDataRange();
    var notes = range.getNotes();
    for (var i = 0; i < notes.length; i++) {
        for (var j = 0; j < notes[i].length; j++) {
            var cellNotes = notes[i][j];
            if (cellNotes !== "") {
                reportSheet.getRange(i + 1, j + 1).clearNote();
            }
        }
    }

    // Проверяем, есть ли скрытые строки
    var maxRows = reportSheet.getLastRow();
    var hiddenRows = [];
    for (var i = 1; i <= maxRows; i++) {
        if (reportSheet.isRowHiddenByUser(i)) {
            hiddenRows.push(i);
        }
    }

    // Если есть скрытые строки, показываем их
    if (hiddenRows.length > 0) {
        reportSheet.showRows(1, maxRows);
    }

    // Если есть активный фильтр, удаляем его
    var filter = reportSheet.getFilter();
    if (filter) {
        filter.remove();
    }

    // надо удалить все, начиная с третьей строки

    if(reportSheet && reportSheet.getLastRow() > 0) {
        reportSheet.getRange(6, 1, reportSheet.getLastRow(), reportSheet.getLastColumn()).clear();
        reportSheet.getRange(6, 1, reportSheet.getLastRow(), reportSheet.getLastColumn()).clearContent();
    }

    reportSheet.getRange('B5').setValue( reportName + ' ' + mondayString + ' - ' + sundayString).setFontSize(20);

    /////////////////////////
    // Вставляем заголовки //
    /////////////////////////

    // PAID HOURS REPORT NOTIFICATION SAMPLE

    // по ходу вывода подробного табличного отчета, нам нужно сформировать короткое текстовое уведомление
    // для добавления в таблицу PaidHoursNotifications (Date - последняя дата данных в отчете - вчерашний день, Status - new или sent, CurrentTime (дата и время формирования отчета),	Notification)
    // отчет нужно формировать только, если у нас weeks = 1.
    // в понедельник мы формируем отчет только если lastweek = true
    // вторник - пятница формируем отчет только если lastweek = false
    // отчет за текущую дату не должен дублироваться, а только обновляться в том же поле за ту же дату

    // в отчете нам нужны следующие данные которые мы соберем по ходу.
    // сначала сформируем заголовок для текущей или прошлой недели
    // для каждого дня в заголовках формируемой

    // let paidHoursDailyReportTitle;
    // например
    // Платные часы (текущая неделя)
    // План 22-28 апреля - 498

    // собираем в объект даты и количество платных часов за день (все dailyHours за день для всех разработчиков)

    // let paidHoursDailyReportListByDate = { date, hours };
    // например
    // 22.04 - 455
    // 23.04 - 474
    // 24.04 - 467
    // 25.04 - 476
    // 26.04 - 417
    // 27.04 - 102
    // 28.04 - 16

    // во второй объект

    // берем данные по отсутствующим часам за последний отчетный день (те у кого выводится темно красный цвет)

    // let paidHoursDailyReportYesterdayNoHours = { developer, hours };
    // например
    // Нет часов
    // Крушеницкий Владислав	0
    // Раткин Иван	0
    // Рубцов Илья	0
    // Цугаев Ахмад	0

    // берем данные по мало платных за последний рабочий день (те у кого выводится светло-желтый), мы можем проверить какой working day был последний на этой неделе

    // let paidHoursDailyReportYesterdayLessHours = { developer, hours };
    // например
    // Мало платных
    // Гордеев Олег	0
    // Петров Егор	4
    // Тумаков Андрей	0

    // формируем выводы если это lastweek, который отправляем в понедельник
    // нам нужно вывести сумму всех платных часов за день dailyHours для все дней недели и всех разработчиков.
    // также берем сумму платных часов по плану, считаем на сколько процент выполнен план
    // берем сумму всех часов - платных и бесплатных без отпуска - то что выводим в четвертой строке - сумма за все дни
    // в конце делим сумму платных на сумму всех часов и получаем процент платных, также в конце выводим. округляем до двух знаков.

    // let paidHoursDailyReportLastWeekStats;
    // например
    // За последнюю неделю
    // Платных 2406.43 (по плану 2472 - выполнен на 97%)
    // Всего 2838.46
    // Процент 84.78%

    // в самом конце все данные складываем в текстовое сообщение, которое кладем в Notification, либо обновляем, если там уже есть данные.

    // отправку добавим позже


    console.log('Вставляем заголовки');

    let column = 2;
    let row = 7;

    reportSheet.getRange(row,column).setValue('Name').setVerticalAlignment("middle").setHorizontalAlignment("center").setFontSize(10);
    column++;
    reportSheet.getRange(row,column).setValue('Plan').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(column, 150);
    column++;
    reportSheet.getRange(row,column).setValue('Fact').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(column, 150);
    column++;

    let endColumn = column;
    let dateRange = getDateRange(mondayDate, sundayDate);
    let weekNumber = weeks;
    let dayNumber = 0;
    let plannedPaidHours = 0;
    let weekPlan = 0;

    // Добавление заголовков для каждой даты
    dateRange.forEach(date => {
        let formattedDate = formatDateForHeader(date);
        let isWorkingDay = isWorkDay(date, allDailyData);
        dayNumber++;

        if (weeks === 1) {
            weekPlan = paidHoursPlan;
            let workingDays = workingHours / 8; // Предполагаем, что рабочий день длится 8 часов

            if (isWorkingDay) {
                plannedPaidHours = weekPlan / workingDays;
                reportSheet.getRange(row, column).setValue(formattedDate).setNote(`Planned paid hours: ${plannedPaidHours.toFixed(2)}`);
            } else {
                reportSheet.getRange(row, column).setValue(formattedDate);
            }
        } else {
            //let weekNumber = getWeekNumberForDate(date, mondayDate, weeks);

            if (weekPlans[weekNumber]) {
                weekPlan = weekPlans[weekNumber].paidHoursPlan || 0;
                let weekWorkingHours = weekPlans[weekNumber].workingHours || 0;
                let workingDays = weekWorkingHours / 8;

                if (isWorkingDay) {
                    let plannedPaidHours = weekPlan / workingDays;
                    reportSheet.getRange(row, column).setValue(formattedDate).setNote(`Planned paid hours: ${plannedPaidHours.toFixed(2)}`);
                } else {
                    reportSheet.getRange(row, column).setValue(formattedDate);
                }
            } else {
                reportSheet.getRange(row, column).setValue(formattedDate);
            }

        }

        if (dayNumber === 7) {
            reportSheet.getRange(row, column).setNote(`Planned week paid hours: ${weekPlan.toFixed(2)}`);
            dayNumber = 0;
            weekNumber--;
        }

        if (!isWorkingDay) reportSheet.getRange(row, column).setBackground(lightRed);

        column++;
    });



    // Write TOTAL in the next two columns
    reportSheet.getRange(row, column)
        .setValue('TOTAL plan')
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setTextRotation(90)
        .setBackground("#ffffff")
        .setFontSize(9)
        .setNote('Paid hours plan + vacation');
    reportSheet.setColumnWidth(column, 50);
    column++;

    // Leave a column for 'fact' data
    reportSheet.getRange(row, column)
        .setValue('TOTAL fact')
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setTextRotation(90)
        .setBackground("#ffffff")
        .setFontSize(9)
        .setNote('Paid hours fact + vacation');
    reportSheet.setColumnWidth(column, 50);
    column++

    reportSheet.getRange(row, column)
        .setValue('Diff')
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setTextRotation(90)
        .setBackground("#ffffff")
        .setFontSize(9)
        .setNote('Difference of Paid hours plan and fact');
    reportSheet.setColumnWidth(column, 50);

    // Add a border to the right of the empty column
    reportSheet.getRange(row, column, 120).setBorder(false, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    column++


    reportSheet.getRange(row, column)
        .setValue('Paid hours')
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setTextRotation(90)
        .setBackground("#ffffff")
        .setFontSize(9)
        .setNote('Paid hours fact without vacation');
    reportSheet.setColumnWidth(column, 60);
    column++


    reportSheet.getRange(row, column)
        .setValue('All hours')
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setTextRotation(90)
        .setBackground("#ffffff")
        .setFontSize(9)
        .setNote('All hours fact (paid and free) without vacation');
    reportSheet.setColumnWidth(column, 60);
    column++


    reportSheet.getRange(row, column)
        .setValue('%')
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setTextRotation(90)
        .setBackground("#ffffff")
        .setFontSize(12)
        .setNote('Percentage of paid hours');
    reportSheet.setColumnWidth(column, 60);

    // Add a border to the right of the empty column
    reportSheet.getRange(row, column, 120).setBorder(false, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    column++
    reportSheet.setColumnWidth(column, 60);
    //const scrumTotalColumn = column;

    column = column++;

    // UNCOMMENT TO DISPLAY PROJETS
    //reportSheet.getRange(3, scrumTotalColumn).setValue(sumColumnE);

    // for (let project of projects) {
    //     // Write the project name and PM initials in the next two columns
    //     reportSheet.getRange(row, column)
    //         .setValue(project.pmInitials + ' ' + project.projectName + ' plan')
    //         .setVerticalAlignment("middle")
    //         .setHorizontalAlignment("center")
    //         .setTextRotation(90)
    //         .setBackground("#cccccc")
    //         .setFontSize(9);
    //     reportSheet.setColumnWidth(column, 40);

    //     // Leave a column for 'fact' data
    //     reportSheet.getRange(row, column + 1)
    //         .setValue(project.pmInitials + ' ' + project.projectName + ' fact')
    //         .setVerticalAlignment("middle")
    //         .setHorizontalAlignment("center")
    //         .setTextRotation(90)
    //         .setBackground("#cccccc")
    //         .setFontSize(9);
    //     reportSheet.setColumnWidth(column + 1, 40);

    //     // Skip an empty column
    //     reportSheet.getRange(row, column + 2).setBackground("#ffffff");
    //     reportSheet.setColumnWidth(column + 2, 40);

    //     // Add a border to the right of the empty column
    //     reportSheet.getRange(row, column + 2, 120).setBorder(false, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    //     // Increment the column counter to skip the 'fact' column
    //     column += 3;
    // }

    /// Вставляем данные
    // UNCOMMENT TO DISPLAY PROJETS

    row++;

    //////////////////
    // ВЫВОД ДАННЫХ //
    //////////////////

    console.log('ВЫВОД ДАННЫХ');

    let totalHoursForDay = [];

    // Объявляем переменные на верхнем уровне функции или цикла
    let weekDevelopers, weekProjects, weekPaidHoursPlan;

    let paidHoursDailyReportTitle = `Платные часы (${isLastWeek ? 'последняя неделя' : 'текущая неделя'})`;
    let paidHoursDailyReportByDate = {};
    let paidHoursDailyReportYesterdayNoHours = [];
    let paidHoursDailyReportYesterdayLessHours = [];
    let totalPaidHours = 0;
    let totalPlannedPaidHours = 0;
    let totalAllHours = 0;
    let lastWorkingDay = null;


    if (weeks === 1) {
        // Если всего одна неделя, используем общие переменные developers и projects
        weekDevelopers = developers;
        weekProjects = projects;
        weekPaidHoursPlan = paidHoursPlan;
    }


    for (let developer of developers) {
        if(!developer.name) continue;
        let developerName = developer.name.split("(")[0].trim(); // Remove everything after the "(" and trim spaces

        let allocationList = '';
        let totalDeveloperHours = 0;
        let totalDeveloperPaidHours = 0;
        if(allAllocationData[developerName] && allAllocationData[developerName].list) allocationList = allAllocationData[developerName].list;

        let developerAllocationData = allAllocationData[developerName];

        // надо построить projectHours из объекта planDevelopers
        let projectHours = '';
        //let projectHoursTotal = 0;
        let projectNames = [];
        for (let planDeveloper of planDevelopers) {
            if (planDeveloper.name === developerName) {
                for (let project in planDeveloper.projects) {
                    let projectObject = projects.find(proj => proj.projectName === project);
                    if (!projectObject) continue;
                    else pmInitials = projectObject.pmInitials;
                    projectNames.push(projectObject.projectName);
                    projectHours += pmInitials + ' ' + project + ' (' + planDeveloper.projects[project] + ')\n';
                    //if (projectObject.projectName !== 'vacation') projectHoursTotal += planDeveloper.projects[project];
                }
                break;
            }
        }

        if (projectHours.length === 0 && allocationList.length === 0) continue;

        row++;
        column=3;
        reportSheet.getRange(row, 2).setValue(developerName).setVerticalAlignment("middle").setWrap(true);
        reportSheet.getRange(row, column).setValue(projectHours).setVerticalAlignment("top").setWrap(true).setFontSize(8);
        reportSheet.getRange(row, column+1).setValue(allocationList).setVerticalAlignment("top").setWrap(true).setFontSize(8);

        column = column + 2;

        //let startDate = getLastMonday(weeks); // получаем дату понедельника для самой старшей недели

        // Подсчет weeklyHours
        let weeklyHoursByWeekNumber = {}; // Словарь для хранения часов по номерам недель
        let currentWeekNumber = weeks; // Начинаем с самой поздней недели и идем назад
        let dayNumber = 0;  // счётчик текущего дня недели

        // Предварительный проход
        dateRange.forEach((date, index) => {
            if (dayNumber === 7) {
                currentWeekNumber--;
                dayNumber = 0; // Сброс счетчика дней при переходе на новую неделю
            }

            let dayHours = 0;
            let dailyDataForAllDevelopers = allDailyData[formatDate(date)];

            // Подсчитываем часы всех разработчиков за день
            if (dailyDataForAllDevelopers) {
                //for (let developerName in dailyDataForAllDevelopers) {
                let developerData = dailyDataForAllDevelopers[developerName];
                if (developerData && developerData.projects) {
                    for (let projectName in developerData.projects) {
                        dayHours += developerData.projects[projectName];
                    }
                }
                //}
            }

            // Накапливаем часы для текущей недели
            if (!weeklyHoursByWeekNumber[currentWeekNumber]) {
                weeklyHoursByWeekNumber[currentWeekNumber] = 0;
            }
            weeklyHoursByWeekNumber[currentWeekNumber] += dayHours;

            dayNumber++;
            if (index === dateRange.length - 1 && dayNumber < 7) { // Обработка последней, возможно неполной, недели
                currentWeekNumber--;
            }
        });

        dayNumber = 0;
        let weekNumber = weeks; // начинаем с самой старшей недели

        dateRange.forEach((date, index) => {
            dayNumber++;

            let weeklyHours = weeklyHoursByWeekNumber[weekNumber];

            // Создаем уведомление для дня
            let formattedDate = formatDate(date);
            let dailyTotalHours = totalHoursForDay[formattedDate] || 0;

            if (!paidHoursDailyReportByDate[formattedDate]) {
                paidHoursDailyReportByDate[formattedDate] = 0;
            }



            if (weeks > 1) {
                // Если недель больше одной, выбираем данные из weekPlans для текущей недели
                let planData = weekPlans[weekNumber] || { developers: [], projects: [], additionalData: [] };
                weekDevelopers = planData.developers;
                weekProjects = planData.projects;

                // Проверка наличия additionalData
                if (planData.additionalData && planData.additionalData.length > 0) {
                    weekPaidHoursPlan = planData.additionalData[0].paidHoursPlan || 0;
                    let weekWorkingHours = planData.additionalData[0].workingHours || 0;
                    let workingDays = weekWorkingHours / 8;
                } else {
                    weekPaidHoursPlan = 0;
                    let workingDays = 0;
                }
            } else {
                let workingDays = workingHours / 8;
            }

            // Массив для хранения общего количества часов всех разработчиков
            let totalHoursPerDeveloper = [];
            let projectHoursTotal = 0;
            let projectNames = [];
            let projectNamesWithFree = [];

            if(weekDevelopers && weekDevelopers.length>0) {
                for (let planDeveloper of weekDevelopers) {
                    let projectHoursTotal = 0; // Общее количество часов для конкретного разработчика

                    for (let project in planDeveloper.projects) {
                        projectHoursTotal += planDeveloper.projects[project]; // Суммируем часы по проектам
                    }
                    // Добавляем общее количество часов в массив
                    totalHoursPerDeveloper.push(projectHoursTotal);
                }
                for (let planDeveloper of weekDevelopers) {
                    if (planDeveloper.name === developerName) {
                        for (let project in planDeveloper.projects) {
                            projectNamesWithFree.push(project);
                            let projectObject = projects.find(proj => proj.projectName === project);
                            if (!projectObject) continue;
                            else pmInitials = projectObject.pmInitials;
                            projectNames.push(projectObject.projectName);
                            //projectHours += pmInitials + ' ' + project + ' (' + planDeveloper.projects[project] + ')\n';
                            if (projectObject.projectName !== 'vacation') projectHoursTotal += planDeveloper.projects[project];
                        }
                        break;
                    }
                }

            }



            let commonTotalHours = findMode(totalHoursPerDeveloper);
            let workingDays = commonTotalHours.map(hours => hours / 8);  // Предполагаем, что каждый рабочий день длится 8 часов
            let dailyData = {};
            let dailyList = '';
            let dailyHours = 0;
            let isWorkingDay = isWorkDay(date, allDailyData);

            if (isWorkingDay) lastWorkingDay = date;

            if (!totalHoursForDay[formatDate(date)]) totalHoursForDay[formatDate(date)] = 0;
            if (allDailyData[formatDate(date)] && allDailyData[formatDate(date)][developerName]) {
                dailyData = allDailyData[formatDate(date)][developerName];
            }


            //if(developerName === 'Сотников Геннадий') console.log(formatDate(date), dailyData, weekProjects.length);

            if (weekProjects && weekProjects.length > 0) {
                // Проходим по проектам текущей недели, не глобальным
                for (let project of weekProjects) {
                    if (dailyData && dailyData.projects && dailyData.projects[project.projectName]) {
                        dailyHours += dailyData.projects[project.projectName];
                        totalDeveloperPaidHours += dailyData.projects[project.projectName];
                    }
                }
            }

            if (dailyData && dailyData.totalHours) {
                dailyList = dailyData.list;
                totalHoursForDay[formatDate(date)] += Number(dailyData.totalHours);
                totalDeveloperHours += Number(dailyData.totalHours);
            }

            reportSheet.getRange(row, column).setValue(dailyHours).setVerticalAlignment("top").setWrap(true).setFontSize(8).setNote(dailyList);
            if (dayNumber === 7) {
                reportSheet.getRange(row, column).setBorder(false, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
                currentDay = 0;
            }

            // Проверяем, присутствуют ли в массиве проекты помимо "vacation"
            const hasVacation = projectNames.some(project => project === "vacation");

            if (!dailyData.list && dailyHours === 0) {
                if (!hasVacation && allDailyData[formatDate(date)] && Object.keys(allDailyData[formatDate(date)]).length>20 && projectNamesWithFree && projectNamesWithFree.length>0 && projectHoursTotal && projectHoursTotal >= 20) {
                    if(isWorkingDay) {
                        addNoHoursNotification(date, developerName, 'Please put the hours for ' + formatDate(date), Object.keys(allDailyData[formatDate(date)]).length>20);
                        paidHoursDailyReportYesterdayNoHours.push({ developer: developer.name, hours: dailyHours, date: date });
                    }
                    // Применяем более темный красный цвет, если есть проекты помимо "vacation"
                    reportSheet.getRange(row, column).setBackground(darkRed);
                } else {
                    // Применяем светлый красный цвет, если в списке только "vacation" и в других случаях (выходные дни)
                    reportSheet.getRange(row, column).setBackground(lightRed);
                }
            } else if (checkNegativeDeviation(projectHoursTotal, dailyHours, workingDays) && checkNegativeDeviationWeek(projectHoursTotal, weeklyHours)) {
                // Применяем светло-желтый цвет, если есть отклонение более 25%
                if(isWorkingDay) paidHoursDailyReportYesterdayLessHours.push({ developer: developer.name, hours: dailyHours, date: date });
                reportSheet.getRange(row, column).setBackground(lightYellow);
            }

            paidHoursDailyReportByDate[formattedDate] += dailyHours;
            totalPlannedPaidHours = weekPaidHoursPlan;




            if (dayNumber === 7) {
                dayNumber = 0;
                weekNumber--;
            }

            column++;
        });


        totalAllHours += totalDeveloperHours;
        totalPaidHours += totalDeveloperPaidHours;

        let columnPlanHours = column;
        let columnFactHours = column+1;
        let columnDiffHours = column+2;
        let planHoursTotal = 0;
        let factHoursTotal = 0;

        column = column+3;

        for (let project of projects) {
            //let dataRow = [];

            // Write the plan hours for the developer for this project
            let planHours = Object.keys(project.developers).find(devName => devName.startsWith(developer.name));
            planHours = project.developers[planHours] || '';


            // Write the fact hours for the developer for this project
            let factHours = '';
            if (developerAllocationData && developerAllocationData.projects) {
                factHours = developerAllocationData.projects[project.projectName] || '';
            }


            // UNCOMMENT TO DISPLAY PROJETS
            // dataRow.push({
            //     value: planHours,
            //     verticalAlignment: "middle",
            //     horizontalAlignment: "center",
            //     background: "#cccccc",
            //     fontSize: 8
            // });

            // dataRow.push({
            //     value: factHours,
            //     verticalAlignment: "middle",
            //     horizontalAlignment: "center",
            //     background: "#cccccc",
            //     fontSize: 8
            // });

            // // Calculate the difference (plan - fact) and write in the next column
            // let formula = `=IF(AND(ISBLANK(R${row}C${column}), ISBLANK(R${row}C${column+1})), "", R${row}C${column+1}-R${row}C${column})`;
            // let difference = (factHours - planHours) || '';
            // let color = difference < 0 ? "red" : "green";
            // dataRow.push({
            //     value: difference,
            //     formula: formula,
            //     verticalAlignment: "middle",
            //     horizontalAlignment: "center",
            //     background: "#ffffff",
            //     fontSize: 8,
            //     fontColor: color
            // });

            // Write all the data to the row at once
            // let range = reportSheet.getRange(row, column, 1, 3);
            // range.setValues([dataRow.map(cell => cell.value)]);
            // range.setBackgrounds([dataRow.map(cell => cell.background)]);
            // range.setFontColors([dataRow.map(cell => cell.fontColor)]);
            // range.setFontSizes([dataRow.map(cell => cell.fontSize)]);
            // range.setVerticalAlignments([dataRow.map(cell => cell.verticalAlignment)]);
            // range.setHorizontalAlignments([dataRow.map(cell => cell.horizontalAlignment)]);

            // UNCOMMENT TO DISPLAY PROJETS


            // Skip an empty column
            reportSheet.getRange(6, column + 2).setBackground("#ffffff");
            reportSheet.setColumnWidth(column + 2, 35);

            // Increment the column counter to skip the 'fact' column
            column += 3;

            planHours = Math.round(planHours * 100) / 100;
            factHours = Math.round(factHours * 100) / 100;

            planHoursTotal += planHours;
            factHoursTotal += factHours;

        }

        if (weeks > 1) {
            // надо добавить данные по отпускам
            let devData = planDevelopers.find(dev => dev.name === developerName);
            let vacationHours = devData.projects['vacation'];
            if (vacationHours !== undefined && vacationHours !== null && vacationHours !== '' && vacationHours !== 0) {
                factHoursTotal += vacationHours;
            }
        }

        let diffHoursTotal = factHoursTotal-planHoursTotal;

        let diffFontColor = "green"
        if(diffHoursTotal < 0) diffFontColor = "red";

        reportSheet.getRange(row, columnPlanHours).setValue(planHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center").setNote(projectHours);
        reportSheet.getRange(row, columnFactHours).setValue(factHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center").setNote(allocationList);
        reportSheet.getRange(row, columnDiffHours).setValue(diffHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center").setFontColor(diffFontColor);

        column = columnDiffHours+1;
        reportSheet.getRange(row, column).setValue(totalDeveloperPaidHours).setVerticalAlignment("middle").setHorizontalAlignment("center");
        column++;
        reportSheet.getRange(row, column).setValue(totalDeveloperHours).setVerticalAlignment("middle").setHorizontalAlignment("center");
        column++;
        reportSheet.getRange(row, column).setValue(totalDeveloperPaidHours/totalDeveloperHours*100+'%').setVerticalAlignment("middle").setHorizontalAlignment("center");



    }

    // Предположим, что в 7 строке указаны даты, и у нас есть totalHoursForDay с рассчитанными суммами
    let dateCellsRange = reportSheet.getRange(7, 2, 1, reportSheet.getLastColumn() - 1);
    let dateValues = dateCellsRange.getValues()[0];

    // Проходим по всем ячейкам с датами
    dateValues.forEach((dateValue, index) => {
        if (dateValue) { // Убедимся, что ячейка не пустая
            // dateValue тут в формате  function formatDateForHeader(date) {
            //     return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd MMM');
            // } нам надо его преобразовать в дату из строки

            let date = new Date(dateValue);
            let dateText = formatDate(date); // Форматируем дату для сопоставления с ключами в totalHoursForDay
            let totalHours = totalHoursForDay[dateText]; // Получаем сумму часов для даты
            if (totalHours !== undefined) {
                // Выводим сумму во второй строке для соответствующей колонки
                totalHours = totalHours.toFixed(2);
                reportSheet.getRange(4, 2 + index).setValue(totalHours);
            }
        }
    });


    let lastColumn = reportSheet.getLastColumn();
    let remainingStartColumn = endColumn;  // Adding 3 to account for fact and plan columns and the column where the next set of data begins
    let totalRow = reportSheet.getLastRow() + 1;

    // Apply the second formula from remainingStartColumn to the last column
    let counter = 5;
    for(let i = remainingStartColumn; i <= lastColumn; i++) {
        let columnLetter = getColumnLetter(i);
        let formula = `=SUM(K9:K${totalRow-1})`;
        formula = formula.replace('K9:K', `${columnLetter}9:${columnLetter}`);

        let cell = reportSheet.getRange(6, i);
        cell.setFormula(formula).setFontSize(8).setHorizontalAlignment("center");

        // Make sure the formula is evaluated
        SpreadsheetApp.flush();

        counter++;

        // If it's the third cell, get its value and set the color accordingly
        if (counter % 3 == 0) {
            let cellValue = cell.getValue();
            if (cellValue < 0) {
                cell.setFontColor("red");
            } else {
                cell.setFontColor("green");
            }
        }
    }

    var filter = reportSheet.getFilter();
    if (filter) {
        filter.remove();
    }

    // создаем новый фильтр
    var range = reportSheet.getRange(8, 2, reportSheet.getLastRow() - 7, lastColumn-1);
    if (range) range.createFilter();

    // Add date and time of data gathering
    const currentTime = new Date().toLocaleString("en-GB", {timeZone: "Asia/Tbilisi"});
    reportSheet.getRange("B6").setValue(`Generated at ${currentTime} (Tbilisi, Georgia Timezone)`);

    if (weeks === 1) {
        totalAllHours = totalAllHours.toFixed(2);
        totalPaidHours = totalPaidHours.toFixed(2);

        paidHoursDailyReportYesterdayNoHours = paidHoursDailyReportYesterdayNoHours.filter(item => item.date === lastWorkingDay);
        paidHoursDailyReportYesterdayLessHours = paidHoursDailyReportYesterdayLessHours.filter(item => item.date === lastWorkingDay);

        let paidHoursDailyReportYesterdayLessHoursFiltered = paidHoursDailyReportYesterdayLessHours.filter(item => {
            // Проверяем, начинается ли имя разработчика с одного из имен из списка pmList
            return !pmList.some(pmName => item.developer.startsWith(pmName));
        });

        // Формирование текста отчета
        let notificationText = `*${paidHoursDailyReportTitle}*\nПлан ${mondayString}-${sundayString} - ${plannedPaidHours}\n`;

        const today = new Date();
        const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
        const todayFormatted = Utilities.formatDate(today, timeZone, 'yyyy-MM-dd');

        Object.entries(paidHoursDailyReportByDate).forEach(([date, hours]) => {
            // Skip today's date and any future dates
            if (new Date(date) >= new Date(todayFormatted)) {
                return;
            }
            notificationText += `${date} - ${hours.toFixed(2)}\n`;
        });

        // Добавление заголовков для списков, если они не пустые
        if (paidHoursDailyReportYesterdayNoHours.length > 0) {
            notificationText += `\n*Нет часов*\n`;
            paidHoursDailyReportYesterdayNoHours.forEach(item => {
                notificationText += `• _${item.developer}_\t${item.hours}\n`;
            });
        }

        if (paidHoursDailyReportYesterdayLessHoursFiltered.length > 0) {
            notificationText += `\n*Мало платных*\n`;
            paidHoursDailyReportYesterdayLessHoursFiltered.forEach(item => {
                notificationText += `• _${item.developer}_\t${item.hours}\n`;
            });
        }

        const percentageCompleted = ((totalPaidHours / totalPlannedPaidHours) * 100).toFixed(2);
        const percentageTotal = ((totalPaidHours / totalAllHours) * 100).toFixed(2);

        if(isLastWeek) {
            // Формирование итогового текста отчета
            notificationText += `\n*За последнюю неделю*\nПлатных ${totalPaidHours} (по плану ${totalPlannedPaidHours} - выполнен на ${percentageCompleted}%)\nВсего ${totalAllHours}\nПроцент ${percentageTotal}%`;
        }

        const notificationsSheet = ss.getSheetByName('PaidHoursNotifications') || ss.insertSheet('PaidHoursNotifications');
        const formattedCurrentTime = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm:ss');

        // Проверка текущего дня недели
        const dayOfWeek = new Date().getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

        if ((isLastWeek && dayOfWeek === 1) || (!isLastWeek && dayOfWeek >= 2 && dayOfWeek <= 5)) {
            // Проверка и обновление уведомлений за последний рабочий день
            if (paidHoursDailyReportYesterdayNoHours.length > 0 || paidHoursDailyReportYesterdayLessHours.length > 0) {
                updateOrAddPaidHoursNotification(
                    notificationsSheet,
                    lastWorkingDay,
                    'new',
                    formattedCurrentTime,
                    notificationText
                );
            }
        }
    }


}


function isWorkDay(date, allDailyData) {
    let formattedDate = formatDate(date);
    if (allDailyData[formattedDate]) {
        return Object.keys(allDailyData[formattedDate]).length > 20;
    }
    return false;
}

function getWeekNumberForDate(date, startDate, totalWeeks) {
    let diffInDays = Math.floor((date - startDate) / (1000 * 60 * 60 * 24));
    return totalWeeks - Math.floor(diffInDays / 7);
}

function getWorkingDaysInWeek(developers) {
    let totalHours = developers.reduce((sum, developer) => {
        return sum + Object.values(developer.projects).reduce((projectSum, hours) => projectSum + hours, 0);
    }, 0);
    return totalHours / (developers.length * 8);
}


function mergeProjects(project1, project2) {
    let mergedProject = {
        pmInitials: '',
        projectName: project1.projectName,
        projectHours: 0,
        developers: {}
    };

    for (let developerName in project1.developers) {
        mergedProject.developers[developerName] = project1.developers[developerName];
    }

    for (let developerName in project2.developers) {
        if (!mergedProject.developers[developerName]) {
            mergedProject.developers[developerName] = project2.developers[developerName];
        } else {
            mergedProject.developers[developerName] += project2.developers[developerName];
        }
    }

    mergedProject.projectHours = project1.projectHours + project2.projectHours;

    mergedProject.pmInitials = project1.pmInitials || project2.pmInitials;

    return mergedProject;
}

function mergeDevelopers(developer1, developer2) {
    let mergedDeveloper = {
        name: developer1.name,
        location: '',
        projectHours: {},
        projects: {},
        vacationHours: developer1.vacationHours + developer2.vacationHours
    };

    // let developer = {
    //     name: developerName,
    //     location: developerLocation,
    //     projectHours,
    //     projects: {},
    //     vacationHours: workloadData[i][projects.indexOf('vacation') + 1] || 0,  // Add vacation hours
    // };

    for (let projectName in developer1.projects) {
        mergedDeveloper.projects[projectName] = developer1.projects[projectName];
    }

    for (let projectName in developer2.projects) {
        if (!mergedDeveloper.projects[projectName]) {
            mergedDeveloper.projects[projectName] = developer2.projects[projectName];
        } else {
            mergedDeveloper.projects[projectName] += developer2.projects[projectName];
        }
    }

    for (let projectName in developer2.projectHours) {
        if (!mergedDeveloper.projectHours[projectName]) {
            mergedDeveloper.projectHours[projectName] = developer2.projectHours[projectName];
        } else {
            mergedDeveloper.projectHours[projectName] += developer2.projectHours[projectName];
        }
    }
    //mergedDeveloper.projectHours = developer1.projectHours + developer2.projectHours;

    return mergedDeveloper;
}



// Этот скрипт сначала вычисляет дневной план по вашему проекту, разделив общее количество часов на 5.
// Затем он проверяет, отклоняется ли количество часов, затраченное в конкретный день, более чем на 25% от этого плана.
// Если отклонение больше 25%, функция вернёт true, в противном случае — false.
function checkDeviation(projectHoursTotal, dailyHours, workingDays) {
    // Считаем дневной план
    const dailyPlan = projectHoursTotal / workingDays;

    // Проверяем, отклоняется ли dailyHours от dailyPlan более чем на 25%
    const deviation = Math.abs(dailyHours - dailyPlan);
    if (deviation > dailyPlan * 0.25) {
        return true; // Отклонение более 25%
    } else {
        return false; // Отклонение 25% или меньше
    }
}

function checkNegativeDeviation(projectHoursTotal, dailyHours, workingDays) {
    const dailyPlan = projectHoursTotal / workingDays;
    const deviation = dailyHours - dailyPlan;

    // Проверяем, что отклонение отрицательное и более чем на 25%
    if (deviation < 0 && Math.abs(deviation) > dailyPlan * 0.25) {
        return true; // Отклонение более 25% и меньше плана
    } else {
        return false; // Отклонение меньше 25% или равно/больше плана
    }
}

function checkNegativeDeviationWeek(projectHoursTotal, weeklyHours) {
    const deviation = weeklyHours - projectHoursTotal;

    // Проверяем, что отклонение отрицательное и более чем на 10%
    if (deviation < 0 && Math.abs(deviation) > projectHoursTotal * 0.1) {
        return true; // Отклонение более 10% и меньше плана
    } else {
        return false; // Отклонение меньше 10% или равно/больше плана
    }
}


// Функция, преобразующая номер столбца в букву
function getColumnLetter(columnNumber) {
    let temp, letter = '';
    while (columnNumber > 0) {
        temp = (columnNumber - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        columnNumber = (columnNumber - temp - 1) / 26;
    }
    return letter;
}

function getAllDevelopersStackDataFromSheet() {
    const developersSheetId = "1VW615PcoaR90HLDD-JQeDmeAcz6DH1T_gCuN17v9C1I";
    const developersSheet = SpreadsheetApp.openById(developersSheetId);
    const namesSheet = developersSheet.getSheetByName("Developers english vs russian names");
    let namesData = namesSheet.getRange(2, 1, namesSheet.getLastRow(), 2).getValues();

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeveloperStackData');
    let data = sheet.getRange(2, 1, sheet.getLastRow(), 6).getValues();

    let allDevelopersData = {};

    namesData.forEach(nameRow => {
        let developerNameRussian = nameRow[1];
        let developerName = nameRow[0]; // Получить имя на английском

        let developerData = data.filter(row => row[0] === developerName);
        let stackData = {};

        developerData.forEach(row => {
            let stack = row[2]; // Технология
            let level = row[3]; // Уровень

            // Подобно тому, как вы обрабатывали уровни в предыдущей функции
            if (level) {
                level = level.toLowerCase().trim().replace('middle', 'mid').replace('junior', 'jun').replace('senior', 'sr').replace('?', '').replace('nonchecked', '');
            }
            if (level != 'nonchecked') stackData[stack] = level || '';
        });

        allDevelopersData[developerNameRussian] = stackData;
    });

    return allDevelopersData;
}

function getResumeLink(developerName) {
    const candidateWorkflowId = "189YZ_AKtBhVBADGksYIjKQCg8h_ky6Bh5tjEzxUWeXY";
    const candidateSpreadsheet = SpreadsheetApp.openById(candidateWorkflowId);
    const candidateSheet = candidateSpreadsheet.getSheetByName('Candidates');

    const lastRow = candidateSheet.getLastRow();
    const candidateDataRange = candidateSheet.getRange(1, 1, lastRow, 12);  // Retrieve the data range including column L (resume link)
    const candidateData = candidateDataRange.getValues();
    const candidateRichTextData = candidateDataRange.getRichTextValues();

    for(let i = lastRow - 1; i >= 0; i--) {  // Start from the bottom row
        if(candidateData[i][0] === developerName) {  // The second column contains the developer's name
            let linkRichText = candidateRichTextData[i][11];  // Get the RichTextValue from column L
            if (linkRichText) {
                let linkUrl = linkRichText.getLinkUrl();
                if (linkUrl) {
                    return linkUrl;  // Return the URL of the resume link
                }
            }
        }
    }
    return "";  // If no matching developer is found or if there is no URL, return an empty string
}


function getAllocationData(developers, projects, isLastWeek = false, isPaidHours = false, weeks) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let scrumSheetName = '';
    let startDay, endDay;
    if (isLastWeek) {
        if (weeks > 1) {
            scrumSheetName = `Scrum files 2024`;
        } else {
            scrumSheetName = `Scrum files for last week`;
        }
    } else {
        scrumSheetName = `Scrum files for current week`;
    }

    const scrumSheet = ss.getSheetByName(scrumSheetName);

    let allocationData = {};
    let dailyData = {};
    let scrumData;

    try {
        const range = scrumSheet.getRange('A3:E' + scrumSheet.getLastRow());
        scrumData = range.getValues();
        // надо почистить scrumDatа от лишних дат, если weeks > 1 И если isLastWeek = true И startDay и endDay не null

        if (weeks && weeks > 1) {
            startDay = getLastMonday(weeks);
            endDay = getLastSunday(1);
            scrumData = scrumData.filter(row => {
                if (!row[1]) return false;
                return row[1] >= startDay && row[1] <= endDay;

            });
        }

    } catch (error) {
        // SpreadsheetApp.getUi().alert("Error retrieving data from the scrum sheet: " + error);
        return;
    }

    scrumData.forEach(row => {
        if (row[2] == "HR") row[3] = "HR";
        if (row[2] == "PRESALE") row[3] = "SALES";
        if (row[2] == "Administrative") row[3] = "Administrative";
        if (row[2] == "Testing") row[3] = "Testing";
        if (row[2] == "DevOps") row[3] = "DevOps";

        const [developerShort, date, type, project, hours] = row;

        if (!row) return null;
        if (!(date && date instanceof Date)) return null;

        const rowDateStr = formatDate(date);

        if (!dailyData[rowDateStr]) {
            dailyData[rowDateStr] = {}; // Инициализация словаря для этой даты
        }

        let roundedHours = Math.round(hours * 100) / 100;
        let developerFull = developers.find(developer => developer.name.startsWith(developerShort));
        if (developerFull && developerFull.name) {
            developerFull.name = developerFull.name.split("(")[0].trim();
            if (!allocationData[developerFull.name]) {
                allocationData[developerFull.name] = {projects: {}, list: '', pmInitials: {}};
            }

            let projectData = projects.find(proj => proj.projectName === project);
            if (projectData) {
                allocationData[developerFull.name].pmInitials[project] = projectData.pmInitials;
            }

            if (!allocationData[developerFull.name].projects[project]) {
                allocationData[developerFull.name].projects[project] = 0;
            }

            allocationData[developerFull.name].projects[project] += roundedHours;

            if (!dailyData[rowDateStr][developerFull.name]) {
                dailyData[rowDateStr][developerFull.name] = {projects: {}, list: '', pmInitials: {}};

            }

            if (!dailyData[rowDateStr][developerFull.name].projects[project]) {
                dailyData[rowDateStr][developerFull.name].projects[project] = 0;
            }

            dailyData[rowDateStr][developerFull.name].projects[project] += roundedHours;




            //// Logger.log(developerFull.name + ' ' + project + ' ' + roundedHours);
        } else {
            // Logger.log(developerShort + " отсутствует в списке");

            let projectData = projects.find(proj => proj.projectName === project);

            if (!allocationData[developerShort]) {
                allocationData[developerShort] = {projects: {}, list: '', pmInitials: {}};
            }

            if (projectData) {
                allocationData[developerShort].pmInitials[project] = projectData.pmInitials;
            } else {
                allocationData[developerShort].pmInitials[project] = '';
            }



            if(!allocationData[developerShort].projects[project]) allocationData[developerShort].projects[project] = 0;
            allocationData[developerShort].projects[project] += roundedHours;

            if (!developers.find(dev => dev.name === developerShort)) {
                developers.push({
                    name: developerShort,
                    location: '', // Местоположение неизвестно
                    projectHours: {[project]: 0}, // Добавляем проект с нулевыми часами
                    projects: {[project]: 0}, // Добавляем проект с нулевыми часами
                    vacationHours: 0 // Часы отпуска устанавливаем в 0
                });
            } else {
                // Если разработчик уже существует в списке developers, просто добавляем проект к его списку проектов
                let existingDeveloper = developers.find(dev => dev.name === developerShort);
                existingDeveloper.projectHours[project] = 0; // Добавляем проект с нулевыми часами
                existingDeveloper.projects[project] = 0; // Добавляем проект с нулевыми часами
            }

            // заполняем также dailyData[rowDateStr][developerName]

            if (!dailyData[rowDateStr][developerShort]) {
                dailyData[rowDateStr][developerShort] = {projects: {}, list: '', pmInitials: {}};
            }

            if (projectData) {
                dailyData[rowDateStr][developerShort].pmInitials[project] = projectData.pmInitials;
            } else {
                dailyData[rowDateStr][developerShort].pmInitials[project] = '';
            }
            if(!dailyData[rowDateStr][developerShort].projects[project]) dailyData[rowDateStr][developerShort].projects[project] = 0;

            dailyData[rowDateStr][developerShort].projects[project] += roundedHours;

        }
    });

    for (let developerName in allocationData) {
        let allocationList = new Set();
        let developerData = developers.find(dev => dev.name === developerName);
        if(!developerData) {
            //console.log(developerName + ' проблемный проблемный проблемный проблемный проблемный');
            continue;
        }
        let developerProjects = developerData.projects;

        // 1. Loop for projects present in the plan (developerProjects)
        for (let project in developerProjects) {
            let hours = allocationData[developerName].projects[project];
            if (!hours) {
                let projectData = projects.find(proj => proj.projectName === project);
                if (projectData) {
                    hours = 0;
                    allocationData[developerName].pmInitials[project] = projectData.pmInitials;
                }
            }

            // If project is 'vacation', get hours from the plan
            if (project === 'vacation') {
                hours = developerData.vacationHours;
            }

            if (!isNaN(hours)) {
                hours = Math.round(hours * 100) / 100;
                let roundedHours = hours.toFixed(2);
                let pmInitials = allocationData[developerName].pmInitials[project];
                allocationList.add(`${pmInitials} ${project} (${roundedHours})`);
            }
        }

        // 2. Loop for projects present in allocationData but not in developerProjects
        for (let project in allocationData[developerName].projects) {
            if (!developerProjects[project]) {  // If project is not already in the developerProjects
                let hours = allocationData[developerName].projects[project];
                hours = Math.round(hours * 100) / 100;

                let roundedHours = hours.toFixed(2);
                let projectData = projects.find(proj => proj.projectName === project);
                let pmInitials = projectData ? projectData.pmInitials : '';

                // Skip entry if pmInitials is undefined or hours is NaN
                if (!isNaN(hours)) {
                    allocationList.add(`${pmInitials} ${project} (${roundedHours})`);
                }
            }
        }

        allocationData[developerName].list = Array.from(allocationList).join('\n');
    }


    // 1. Loop for projects present in the plan (developerProjects) and update dailyData
    for (let date in dailyData) {
        for (let developerName in dailyData[date]) {
            if (!dailyData[date][developerName].totalHours) dailyData[date][developerName].totalHours = 0;
            let dailyAllocationList = new Set();
            let developerProjects = dailyData[date][developerName].projects;

            for (let project in developerProjects) {
                let hours = dailyData[date][developerName].projects[project];
                if (!hours) {
                    let projectData = projects.find(proj => proj.projectName === project);
                    if (projectData) {
                        hours = 0;
                        dailyData[date][developerName].pmInitials[project] = projectData.pmInitials;
                    }
                }

                // If project is 'vacation', get hours from the plan
                if (project === 'vacation') {
                    hours = developerData.vacationHours;
                }

                if (!isNaN(hours)) {
                    if (project !== 'vacation') dailyData[date][developerName].totalHours += hours;
                    hours = Math.round(hours * 100) / 100;
                    let roundedHours = hours.toFixed(2);
                    let pmInitials = allocationData[developerName].pmInitials[project];
                    dailyAllocationList.add(`${pmInitials} ${project} (${roundedHours})`);
                }
            }
            dailyData[date][developerName].list = Array.from(dailyAllocationList).join('\n');
            // dailyData[date][developerName].totalHours = Array.from(dailyAllocationList).reduce((acc, curr) => {
            //     let hours = parseFloat(curr.match(/\d+\.\d+/)[0]);
            //     return acc + hours;
            // });
        }
    }

    // Сохраняем списки до пересборки данных
    let savedLists = {};
    if (isPaidHours) {
        savedLists = saveLists(allocationData, dailyData);
    }



    if(isPaidHours) {
        allocationData = {};
        dailyData = {};

        scrumData.forEach(row => {
            if (row[2] == "HR") row[3] = "HR";
            if (row[2] == "PRESALE") row[3] = "SALES";
            if (row[2] == "Administrative") row[3] = "Administrative";
            if (row[2] == "Testing") row[3] = "Testing";
            if (row[2] == "DevOps") row[3] = "DevOps";

            const [developerShort, date, type, project, hours] = row;

            if (!date) return null;
            // if (date && (!date.split('/')[0] || !date.split('/')[1] || !date.split('/')[2])) return null;

            // if (weeks && weeks > 1) {
            //     if (date.getDate() < startDay || date.getDate() > endDay) return null;
            // }

            //if (isPaidHours) {
            if (type !== 'DEV' && type !== 'PM' && project !== 'vacation') {
                return null;
            }
            //}

            const rowDateStr = formatDate(date); // Форматирование даты для ключа словаря
            if (!dailyData[rowDateStr]) {
                dailyData[rowDateStr] = {}; // Инициализация словаря для этой даты
            }

            let roundedHours = Math.round(hours * 100) / 100;
            let developerFull = developers.find(developer => developer.name.startsWith(developerShort));
            if (developerFull && developerFull.name) {
                developerFull.name = developerFull.name.split("(")[0].trim();
                if (!allocationData[developerFull.name]) {
                    allocationData[developerFull.name] = {projects: {}, list: '', pmInitials: {}};
                }

                let projectData = projects.find(proj => proj.projectName === project);
                if (projectData) {
                    allocationData[developerFull.name].pmInitials[project] = projectData.pmInitials;
                }

                if (!allocationData[developerFull.name].projects[project]) {
                    allocationData[developerFull.name].projects[project] = 0;
                }

                allocationData[developerFull.name].projects[project] += roundedHours;

                // заполняем также dailyData[rowDateStr][developerName]

                if (!dailyData[rowDateStr][developerFull.name]) {
                    dailyData[rowDateStr][developerFull.name] = {projects: {}, list: '', pmInitials: {}};

                }

                if (!dailyData[rowDateStr][developerFull.name].projects[project]) {
                    dailyData[rowDateStr][developerFull.name].projects[project] = 0;
                }
                dailyData[rowDateStr][developerFull.name].projects[project] += roundedHours;
                // // Logger.log(developerFull.name + ' ' + project + ' ' + roundedHours);
            } else {
                // Logger.log(developerShort + " отсутствует в списке");
                if (!allocationData[developerShort]) {
                    allocationData[developerShort] = {projects: {}, list: '', pmInitials: {}};
                }
                allocationData[developerShort].pmInitials[project] = '';
                if (!allocationData[developerShort].projects[project]) allocationData[developerShort].projects[project] = 0;
                allocationData[developerShort].projects[project] += roundedHours;

                if (!developers.find(dev => dev.name === developerShort)) {
                    developers.push({
                        name: developerShort,
                        location: '', // Местоположение неизвестно
                        projectHours: {[project]: 0}, // Добавляем проект с нулевыми часами
                        projects: {[project]: 0}, // Добавляем проект с нулевыми часами
                        vacationHours: 0 // Часы отпуска устанавливаем в 0
                    });
                } else {
                    // Если разработчик уже существует в списке developers, просто добавляем проект к его списку проектов
                    let existingDeveloper = developers.find(dev => dev.name === developerShort);
                    existingDeveloper.projectHours[project] = 0; // Добавляем проект с нулевыми часами
                    existingDeveloper.projects[project] = 0; // Добавляем проект с нулевыми часами
                }

                // заполняем также dailyData[rowDateStr][developerName]

                if (!dailyData[rowDateStr][developerShort]) {
                    dailyData[rowDateStr][developerShort] = {projects: {}, list: '', pmInitials: {}};
                }

                dailyData[rowDateStr][developerShort].pmInitials[project] = '';
                if (!dailyData[rowDateStr][developerShort].projects[project]) dailyData[rowDateStr][developerShort].projects[project] = 0;

                dailyData[rowDateStr][developerShort].projects[project] += roundedHours;

            }
        });

        restoreLists(allocationData, dailyData, savedLists);
    }

    return {allocationData, dailyData, scrumData};
}



// Функция для сохранения списков
function saveLists(allocationData, dailyData) {
    let savedLists = {
        allocationData: {},
        dailyData: {}
    };

    for (let developerName in allocationData) {
        savedLists.allocationData[developerName] = allocationData[developerName].list;
    }

    for (let date in dailyData) {
        savedLists.dailyData[date] = {};
        for (let developerName in dailyData[date]) {
            // Инициализация объекта для разработчика в данной дате перед сохранением данных
            if (!savedLists.dailyData[date][developerName]) {
                savedLists.dailyData[date][developerName] = {};
            }
            savedLists.dailyData[date][developerName].list = dailyData[date][developerName].list;
            savedLists.dailyData[date][developerName].totalHours = dailyData[date][developerName].totalHours;
        }
    }

    return savedLists;
}

// Функция для восстановления списков
function restoreLists(allocationData, dailyData, savedLists) {
    for (let developerName in savedLists.allocationData) {
        // Если разработчик уже есть в allocationData, обновляем список
        // Если нет - создаем новую запись
        if (!allocationData[developerName]) {
            allocationData[developerName] = {projects: {}, list: '', pmInitials: {}};
        }
        allocationData[developerName].list = savedLists.allocationData[developerName];
    }

    for (let date in savedLists.dailyData) {
        for (let developerName in savedLists.dailyData[date]) {
            // Если разработчик уже есть в dailyData за эту дату, обновляем список
            // Если нет - создаем новую запись
            if (!dailyData[date][developerName]) {
                dailyData[date][developerName] = {projects: {}, list: '', totalHours: 0, pmInitials: {}};
            }
            dailyData[date][developerName].list = savedLists.dailyData[date][developerName].list;
            dailyData[date][developerName].totalHours = savedLists.dailyData[date][developerName].totalHours;
        }
    }
}


function formatDate(date) {
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0'); // месяцы с 0, поэтому добавляем 1
    const day = date.getDate().toString().padStart(2, '0');
    return `${year}-${month}-${day}`;
}


// function formatDate(date) {
//     // Функция для форматирования даты в строку 'YYYY-MM-DD'
//     return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
// }

function formatDateForHeader(date) {
    const day = date.getDate().toString().padStart(2, '0');
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const month = monthNames[date.getMonth()];
    return `${day} ${month}`;
}


function getDateRange(startDate, endDate) {
    let dateArray = [];
    let currentDate = new Date(startDate);
    while (currentDate <= endDate) {
        dateArray.push(new Date(currentDate));
        currentDate.setDate(currentDate.getDate() + 1);
    }
    return dateArray;
}

function getDevelopers(workloadSheet, all, allocationData, dailyData, isLastWeek = false, isPaidHours = false, weekNumber = 0, workloadSpreadsheet = null) {
    if(!workloadSheet) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const workloadSheetId = "1N65NUtqBA855C6K8swmeFQ9HbvIZU4fq4EnhYzvNV7Q";
        if (!workloadSpreadsheet) workloadSpreadsheet = SpreadsheetApp.openById(workloadSheetId);

        let weekMondayDate;
        let weekSundayDate;

        if (isLastWeek) {
            weekMondayDate = getLastMonday();
            weekSundayDate = getLastSunday();
        } else if (weekNumber < 0) {
            // Для прошлых недель: использовать отрицательные значения weekNumber
            weekMondayDate = getLastMonday(Math.abs(weekNumber));
            weekSundayDate = getLastSunday(Math.abs(weekNumber));
        } else if (weekNumber > 0) {
            // Для будущих недель: использовать положительные значения weekNumber
            weekMondayDate = getNextMonday(weekNumber);
            weekSundayDate = getNextSunday(weekNumber);
        } else {
            // Для текущей недели
            weekMondayDate = getCurrentMonday();
            weekSundayDate = getCurrentSunday();
        }

        const weekMondayString = Utilities.formatDate(weekMondayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();

        const weekSundayString = Utilities.formatDate(weekSundayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();
        let workloadSheetName = weekMondayDate.getMonth() === weekSundayDate.getMonth() ?
            `${weekMondayString.split(" ")[0]}-${weekSundayString.split(" ")[0]} ${weekSundayString.split(" ")[1]}` :
            `${weekMondayString}-${weekSundayString}`;

        if (weekNumber>0) {
            workloadSheetNameNoPlan = workloadSheetName;
            workloadSheetName = workloadSheetName+' (plan)';
        }

        workloadSheet = workloadSpreadsheet.getSheetByName(workloadSheetName);
        if (!workloadSheet) {
            workloadSheet = workloadSpreadsheet.getSheetByName(workloadSheetNameNoPlan);
            if (!workloadSheet) {
                SpreadsheetApp.getUi().alert(`Cannot find sheet "${workloadSheetName}" in the workload spreadsheet.`);
                return [];
            }
        }
    }

    let developers = [];
    let projects = [];
    let isReserve1 = false; // flag to check if we are in the "Запас" section
    let isReserve2 = false;

    let workloadData = workloadSheet.getDataRange().getValues();

    // Retrieve projects from the 5th row
    projects = workloadData[4].slice(1);

    // Iterate through the rows of the workloadData
    for (let i = 5; i < workloadData.length; i++) {
        // Get the developer's name, which is assumed to be in the 4th column
        let developerName = workloadData[i][3];
        let developerLocation = workloadData[i][0];
        developerName = developerName.toString().split("(")[0].trim();

        // Logger.log(developerName);

        const projectHours = getHoursByNameAndProject(workloadData, developerName, allocationData, isPaidHours = false);

        // If the developer name is "total", set the isReserve flag to true
        if (developerName === 'total') {
            if (!isPaidHours) {  //  && !isLastWeek
                isReserve1 = true;
                //console.log('Reserve 1 is set');
            } else {
                break;
            }
        }

        // If we are in the "Запас" section and the developer name is empty, stop the loop
        if (isReserve1 && isReserve2 && !developerName) {
            //console.log('Reserve1 and Reserve 2');
            break;
        }

        // If the developer name is "Запас", continue to the next iteration
        if (developerLocation === 'Запас' && isReserve1) {
            isReserve2 = true;
            //console.log('Reserve 2 is set');
            continue;
        }

        if (isReserve1 && !isReserve2) {
            //console.log('Reserve 1');
            continue;
        }

        // Create a new developer object
        let developer = {
            name: developerName,
            location: developerLocation,
            projectHours: projectHours,
            projects: {},
            vacationHours: workloadData[i][projects.indexOf('vacation') + 1] || 0,  // Add vacation hours
        };

        let workedOnTraining = false;
        let workedOnSales = false;

        for (let j = 5; j < workloadData[i].length; j++) {
            hours = workloadData[i][j] || 0;
            let projectName = projects[j - 1];

            if(!projectName) continue;
            if(projectName === 'd' || projectName === 'pm') continue;

            if (isPaidHours && isPaidHoursProject(projectName)) continue;

            if (hours>0) {
                developer.projects[projectName] = hours;
                if (projectName == "Training") {
                    workedOnTraining = true;
                } else if (projectName == "SALES") {
                    workedOnSales = true;
                }
            }

            if (isReserve2) {
                workedOnTraining = true;
                workedOnSales = true;
            }
        }

        //if (all || (workedOnTraining || workedOnSales)) {
        developers.push(developer);
        //}




    }

    return developers;
}


function getProjects(workloadSheet, projectNameFilter, isLastWeek = false, isPaidHours = false, weekNumber = 0, workloadSpreadsheet = null) {
    let workloadSheetName = '';
    if (!workloadSheet) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const workloadSheetId = "1N65NUtqBA855C6K8swmeFQ9HbvIZU4fq4EnhYzvNV7Q";
        if (!workloadSpreadsheet) workloadSpreadsheet = SpreadsheetApp.openById(workloadSheetId);

        let weekMondayDate;
        let weekSundayDate;

        if (isLastWeek) {
            weekMondayDate = getLastMonday();
            weekSundayDate = getLastSunday();
        } else if (weekNumber < 0) {
            weekMondayDate = getLastMonday(Math.abs(weekNumber));
            weekSundayDate = getLastSunday(Math.abs(weekNumber));
        } else if (weekNumber > 0) {
            weekMondayDate = getNextMonday(weekNumber);
            weekSundayDate = getNextSunday(weekNumber);
        } else {
            weekMondayDate = getCurrentMonday();
            weekSundayDate = getCurrentSunday();
        }

        const weekMondayString = Utilities.formatDate(weekMondayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();
        const weekSundayString = Utilities.formatDate(weekSundayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();

        workloadSheetName = weekMondayDate.getMonth() === weekSundayDate.getMonth() ?
            `${weekMondayString.split(" ")[0]}-${weekSundayString.split(" ")[0]} ${weekSundayString.split(" ")[1]}` :
            `${weekMondayString}-${weekSundayString}`;

        if (weekNumber > 0) {
            workloadSheetNameNoPlan = workloadSheetName;
            workloadSheetName = workloadSheetName + ' (plan)';
        }

        workloadSheet = workloadSpreadsheet.getSheetByName(workloadSheetName);
        if (!workloadSheet) {
            workloadSheet = workloadSpreadsheet.getSheetByName(workloadSheetNameNoPlan);
            if (!workloadSheet) {
                return { projects: [], workloadSheetName: '', additionalData: [] };
            }
        }
    }

    let projects = [];
    let developers = [];

    let workloadData = workloadSheet.getDataRange().getValues();

    developers = workloadData.map(row => row[3]);

    outer: for (let j = 1; j < workloadData[0].length; j++) {
        let projectName = workloadData[4][j].trim() || "(noname)";

        if (isPaidHours && isPaidHoursProject(projectName)) continue;

        let pmInitials = workloadData[0][j].trim();

        let regex = /[а-яё]{2}/i;
        if (pmInitials === undefined || pmInitials === null || pmInitials.length !== 2 || !regex.test(pmInitials)) {
            continue outer;
        }

        if (projectNameFilter && projectName !== projectNameFilter) {
            continue outer;
        }

        let project = { pmInitials, projectName, projectHours: 0, developers: {} };

        for (let i = 5; i < workloadData.length; i++) {
            let hours = workloadData[i][j] || 0;
            let developerName = developers[i];

            developerName = developerName.split("(")[0].trim();

            if (hours >= 0) {
                project.developers[developerName] = hours;
                project.projectHours += hours;
            }
        }

        projects.push(project);
    }

    let additionalData = extractAdditionalData(workloadSheet);

    // Добавляем данные из ячейки A3 в additionalData
    let workingHours = workloadSheet.getRange("A3").getValue();
    additionalData.push({ workingHours: workingHours });

    return { projects, workloadSheetName, additionalData };
}


function extractAdditionalData(sheet) {
    let dataRange = sheet.getDataRange().getValues();
    let additionalData = [];

    for (let i = 0; i < dataRange.length; i++) {
        let row = dataRange[i];
        for (let j = 0; j < row.length; j++) {
            if (row[j] === "часы без отпуска и без бесплатных") {
                let valueColumnIndex = j + 1;
                if (valueColumnIndex < row.length) {
                    let value = row[valueColumnIndex];
                    if (typeof value === 'number') {
                        additionalData.push(value);
                    }
                }
            }
        }
    }
    return additionalData;
}


function getHoursByNameAndProject(data, name, allocationData, isPaidHours = false) {
    var hoursAndProjects = [];
    for (var i = 0; i < data.length; i++) {
        var rowName = data[i][3].toString();
        if (rowName.startsWith(name)) {
            for (var j = 5; j < data[0].length; j++) {
                var cellValue = data[i][j];
                var hours = Math.round(cellValue * 100) / 100;
                if (hours > 0) {
                    var pm = data[0][j];
                    var project = data[4][j];
                    project = project.trim();
                    if (pm == '') break;
                    // если проект 	Testing	Training	Bench	SALES	HR	freelance    Administrative	DevOps	надо пропустить
                    if (isPaidHours && isPaidHoursProject(project)) continue;
                    hoursAndProjects.push(pm + " " + project + " (" + hours.toFixed(2) + ")");
                }
            }
            break;
        }
    }
    if(allocationData) {
        // Проверяем allocationData на наличие проектов, которых не было в плане
        if (allocationData[name]) {
            for (let project in allocationData[name].projects) {
                // Проверяем, существует ли проект с тем же именем в hoursAndProjects
                if (!hoursAndProjects.some(item => item.includes(project))) {
                    hoursAndProjects.push(allocationData[name].pmInitials[project] + " " + project + " (0.00)");  // Добавляем проект с 0 часами
                }
            }
        }
    }
    return hoursAndProjects.join('\n');
}


function isPaidHoursProject(projectName) {
    if (projectName == 'Testing' ||
        projectName == 'Training' ||
        projectName == 'Bench' ||
        projectName == 'SALES' ||
        projectName == 'HR' ||
        projectName == 'freelance' ||
        projectName == 'Administrative' ||
        projectName == 'DevOps' ||
        projectName == 'Site' ||
        projectName == 'Techlead' ||
        projectName == '(noname)'
    ) return true;
    else return false;
}


function getScrumFilesData(fromDate, toDate) {
    const spreadsheetId = "1PntBe9VKwaXDsI-iOk-CibuZ6X1ZFKQNXNOfh76ysBw";
    const sheetName = "Scrum files";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const scrumFilesSheet = ss.getSheetByName(sheetName);

    // Retrieve list of scrum files urls starting from the 3rd row of the 'B' column
    const scrumFilesUrls = scrumFilesSheet.getRange('B3:B' + scrumFilesSheet.getLastRow()).getValues().flat();

    if(scrumFilesUrls.length == 0) {
        // Logger.log('No matching files found.');
        return;
    }

    let data = {};
    for (let url of scrumFilesUrls) {
        if (url) { // check if the cell isn't empty
            const externalFile = SpreadsheetApp.openByUrl(url);
            const monthNames = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
            const monthSheetsToProcess = monthNames.filter(monthName => externalFile.getSheetByName(monthName));

            for (let monthSheetName of monthSheetsToProcess) {
                const externalSheet = externalFile.getSheetByName(monthSheetName);
                const lastRow = externalSheet.getLastRow();
                const monthSheetData = externalSheet.getRange(2, 1, lastRow - 1, 5).getValues();

                monthSheetData.forEach(function(rowData) {
                    if (rowData[0] && rowData[1] && rowData[2] && rowData[4]) {
                        const dateTime = new Date(rowData[0]);
                        if (dateTime >= fromDate && dateTime <= toDate) {
                            const developer = externalFile.getName(); // Here we reference the filename of the external file
                            const dateScrum = Utilities.formatDate(rowData[0], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
                            const typeScrum = rowData[1];
                            const projectScrum = rowData[2];
                            const hoursScrum = rowData[4];
                            if (!data[developer]) {
                                data[developer] = [];
                            }
                            data[developer].push({date: dateScrum, type: typeScrum, project: projectScrum, hours: hoursScrum});
                        }
                    }
                });
            }
        }
    }

    // Logger.log(JSON.stringify(data)); // Log the entire data object
    return data;
}


function getColumnLetter(column) {
    let letter = "";
    while (column > 0) {
        const remainder = (column - 1) % 26;
        letter = String.fromCharCode(65 + remainder) + letter;
        column = Math.floor((column - 1) / 26);
    }
    return letter;
}


function getLastMonday(weeks = 1) {
    const currentMonday = getCurrentMonday();
    const lastMonday = new Date(currentMonday.getFullYear(), currentMonday.getMonth(), currentMonday.getDate() - 7*weeks);
    return lastMonday;
}


function getLastSunday(weeks = 1) {
    const currentSunday = getCurrentSunday();
    const lastSunday = new Date(currentSunday.getFullYear(), currentSunday.getMonth(), currentSunday.getDate() - 7*weeks);
    return lastSunday;
}

function getNextMonday(weeks = 1) {
    const currentMonday = getCurrentMonday();
    const nextMonday = new Date(currentMonday.getFullYear(), currentMonday.getMonth(), currentMonday.getDate() + 7*weeks);
    return nextMonday;
}

function getNextSunday(weeks = 1) {
    const currentSunday = getCurrentSunday();
    const nextSunday = new Date(currentSunday.getFullYear(), currentSunday.getMonth(), currentSunday.getDate() + 7*weeks);
    return nextSunday;
}

function getCurrentMonday() {
    const today = new Date();
    const day = today.getDay();
    const diffToMonday = day === 0 ? -6 : 1 - day; // If today is Sunday (0), we need to subtract 6 to get to the last Monday. Otherwise, subtract the number of days up to Monday
    const currentMonday = new Date(today.getFullYear(), today.getMonth(), today.getDate() + diffToMonday);
    return currentMonday;
}

function getCurrentSunday() {
    const today = new Date();
    const day = today.getDay();
    const diffToNextSunday = day === 0 ? 0 : 7 - day; // Here we calculate the number of days remaining to next Sunday
    const currentSunday = new Date(today.getFullYear(), today.getMonth(), today.getDate() + diffToNextSunday);
    return currentSunday;
}

function gatherDataInSheetCommand() {
    gatherDataInSheet(false);
}

function gatherDataInSheet(isLastWeek = false) {
    // Get Monday and Sunday dates based on isLastWeek
    const startDate = isLastWeek ? getLastMonday() : getCurrentMonday();
    const endDate = isLastWeek ? getLastSunday() : getCurrentSunday();

    const data = getScrumFilesData(startDate, endDate);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Define the sheet name based on isLastWeek
    let sheetName = isLastWeek ? "Scrum files for last week" : "Scrum files for current week";

    let dataSheet = ss.getSheetByName(sheetName);

    // Create new sheet if it doesn't exist
    if (!dataSheet) {
        dataSheet = ss.insertSheet(sheetName);
    } else {
        // Clear the sheet if it already exists
        dataSheet.clear();
    }

    // Add date and time of data gathering
    const currentTime = new Date().toLocaleString("en-GB", {timeZone: "Asia/Tbilisi"});
    dataSheet.getRange("A1").setValue(`Data gathered at ${currentTime} (Tbilisi, Georgia Timezone)`);

    // Headers for the new sheet moved one row down
    const headers = ["Developer", "Date", "Type", "Project", "Hours"];
    dataSheet.getRange(2, 1, 1, headers.length).setValues([headers]);

    let currentRow = 3;

    for (let developer in data) {
        for (let entry of data[developer]) {
            const row = [developer, entry.date, entry.type, entry.project, entry.hours];
            dataSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
            currentRow++;
        }
    }
}

function gatherDataInSheetLastWeek() {
    gatherDataInSheet(true)
}


function gatherScrumFilesDataFromFolder(folderId) {
    if(!folderId) folderId = "1cK4WNQFOHrRIFwPEewpKVtuLkV2b46WH"; // папка 2024
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();
    const files = folder.getFiles();

    const year = parseInt(folderName);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Scrum files " + folderName;

    let scrumFilesSheet = ss.getSheetByName(sheetName);
    if (!scrumFilesSheet) {
        scrumFilesSheet = ss.insertSheet(sheetName);
    }

    // Clear the sheet before filling
    scrumFilesSheet.clear();

    // Create headers
    const headers = ["Developer", "Date", "Type", "Project", "Hours"];
    scrumFilesSheet.getRange(2, 1, 1, headers.length).setValues([headers]);

    while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();
        const externalFile = SpreadsheetApp.openById(file.getId());
        const monthNames = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
        // const monthNames = ["Октябрь"];
        // const monthNames = ["", "", "", "", "", "", "", "", "", "Октябрь", "", ""];
        const monthSheetsToProcess = monthNames.filter(monthName => externalFile.getSheetByName(monthName));

        // Logger.log(`Processing file ${fileName}.`);
        for (let monthSheetName of monthSheetsToProcess) {
            const monthIndex = monthNames.indexOf(monthSheetName) + 1; // Get the month number
            const externalSheet = externalFile.getSheetByName(monthSheetName);
            const lastRow = externalSheet.getLastRow();
            const monthSheetData = externalSheet.getRange(2, 1, lastRow - 1, 5).getValues();
            const preparedData = []; // To collect data for a month

            // Logger.log(`Processing sheet ${monthSheetName} in file ${fileName}.`);
            monthSheetData.forEach(function(rowData) {
                if (rowData[0] && rowData[1] && rowData[2] && rowData[4]) {
                    const dateTime = new Date(rowData[0]);
                    const dateYear = dateTime.getFullYear();
                    const dateMonth = dateTime.getMonth() + 1; // Get the month number

                    if (dateYear === year && dateMonth === monthIndex) {
                        const dateScrum = Utilities.formatDate(rowData[0], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
                        const typeScrum = rowData[1];
                        const projectScrum = rowData[2];
                        const hoursScrum = rowData[4];
                        preparedData.push([fileName, dateScrum, typeScrum, projectScrum, hoursScrum]);
                    } else {
                        // Logger.log(`Data in row does not belong to the year and month. Skipping data.`);
                    }
                }

            });
            // Write all data for the month at once
            if (preparedData.length > 0) {
                // Logger.log('TOTAL Rows in list ' + preparedData.length);
                const startRow = scrumFilesSheet.getLastRow() + 1; // Get the starting row
                scrumFilesSheet.getRange(startRow, 1, preparedData.length, headers.length).setValues(preparedData);
            }
        }
    }

    // Add date and time of data gathering
    const currentTime = new Date().toLocaleString("en-GB", {timeZone: "Asia/Tbilisi"});
    scrumFilesSheet.getRange("A1").setValue(`Data gathered at ${currentTime} (Tbilisi, Georgia Timezone)`);

    // Logger.log(`Data gathering finished for ${sheetName}.`);
}

function collectDeveloperVacationData() {
    // Открытие исходного документа
    var sourceSpreadsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1e2a-p6qW5RfVNsaxMUFKHRYXzUn7b0qrHXjmAKAG7h0/edit');

    // Получение текущего месяца и года на русском языке
    var currentDate = new Date();
    var monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    var currentMonth = monthNames[currentDate.getMonth()];
    var currentYear = currentDate.getFullYear();

    // Получение листа с текущим месяцем
    var currentMonthSheet = sourceSpreadsheet.getSheetByName(currentMonth + ' ' + currentYear);
    if (!currentMonthSheet) {
        // Logger.log('Лист с названием текущего месяца не найден.');
        return;
    }

    // Получение данных с листа текущего месяца
    var dataRange = currentMonthSheet.getDataRange();
    var dataValues = dataRange.getValues();

    // Индексы столбцов "ФИО" и "vacation, remainder"
    var fioIndex = -1;
    var vacationIndex = -1;

    // Находим индексы столбцов "ФИО" и "vacation, remainder"
    for (var i = 0; i < dataValues[0].length; i++) {
        if (dataValues[0][i] === 'ФИО') {
            fioIndex = i;
        } else if (dataValues[0][i] === 'vacation, remainder') {
            vacationIndex = i;
        }
    }

    // Проверка наличия обоих столбцов
    if (fioIndex !== -1 && vacationIndex !== -1) {
        // Фильтрация исходных данных
        var filteredData = dataValues.slice(1).map(function(row) {
            var fio = row[fioIndex];
            var vacation = row[vacationIndex];

            // Проверка наличия значения в обоих столбцах
            if (fio !== '' && vacation !== '') {
                return [fio, vacation];
            }

            return null; // Пропускаем пустые строки
        }).filter(Boolean); // Фильтрация из массива всех значений, равных null

        // Сортировка данных по ФИО (по первому столбцу)
        filteredData.sort(function(a, b) {
            return a[0].localeCompare(b[0]);
        });

        // Создание или открытие листа DeveloperVacation
        var destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        var destinationSheet = destinationSpreadsheet.getSheetByName('DeveloperVacation');
        if (!destinationSheet) {
            destinationSheet = destinationSpreadsheet.insertSheet('DeveloperVacation');
        }

        // Очистка листа DeveloperVacation
        destinationSheet.clear();

        // Записываем время генерации данных
        var now = Utilities.formatDate(new Date(), 'Asia/Tbilisi', 'dd/MM/yyyy, HH:mm:ss');
        destinationSheet.getRange('A1').setValue('Generated at ' + now + ' (Tbilisi, Georgia Timezone)');

        // Запись заголовков второй строкой
        destinationSheet.getRange(2, 1, 1, 2).setValues([['ФИО', 'vacation, remainder']]);

        // Запись отсортированных данных с третьей строки
        destinationSheet.getRange(3, 1, filteredData.length, 2).setValues(filteredData);
    } else {
        // Logger.log('Столбец "ФИО" или "vacation, remainder" не найден на листе текущего месяца.');
        return HtmlService.createHtmlOutput("Столбец 'ФИО' или 'vacation, remainder' не найден на листе текущего месяца.");
    }
}


function collectCandidatesData() {
    var sourceDocId = '189YZ_AKtBhVBADGksYIjKQCg8h_ky6Bh5tjEzxUWeXY';
    var sourceSheetName = 'Candidates';
    var destSheetName = 'CandidatesData';

    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    const developersSheetId = "1VW615PcoaR90HLDD-JQeDmeAcz6DH1T_gCuN17v9C1I";
    const developersSpreadsheet = SpreadsheetApp.openById(developersSheetId);
    const developersSheet = developersSpreadsheet.getSheetByName("Developers english vs russian names");
    var developersRange = developersSheet.getDataRange();
    var developersValues = developersRange.getValues().map(row => row[1]).slice(1);

    var developersNames = developersValues.flat();  // Получаем список разработчиков
    // Logger.log("Developers Names: " + developersNames);

    var sourceSheet = SpreadsheetApp.openById(sourceDocId).getSheetByName(sourceSheetName);
    var sourceRange = sourceSheet.getDataRange();
    var sourceValues = sourceRange.getValues();
    var sourceRichTextValues = sourceRange.getRichTextValues();

    var destSheet = activeSpreadsheet.getSheetByName(destSheetName);
    if (!destSheet) {
        // Если лист не существует, создаем новый
        destSheet = activeSpreadsheet.insertSheet(destSheetName);
    }

    // Очищаем лист
    destSheet.clear();

    // Записываем заголовки
    var headers = sourceValues[0];
    headers.push('ID');
    destSheet.getRange('A2:' + columnToLetter(headers.length) + '2').setValues([headers]);

    // Logger.log("Headers written to destination sheet");

    // Записываем данные кандидатов
    for (var sourceRowIndex = 1; sourceRowIndex < sourceValues.length; sourceRowIndex++) {
        var row = sourceValues[sourceRowIndex];
        var name = row[0];
        var status = row[2];  // Исправлен индекс на 2, так как в колонке C статус

        // Ожидает Решения Резерв

        if (developersNames.includes(name) && (status === 'Принят' || status === 'Ожидает Интервью у клиента' || status === 'Ожидает Решения' || status === 'Резерв' || status === 'Пауза / Думает' || status === 'Пауза / Резерв')) {
            // Logger.log("Found matching candidate: " + name);

            // Добавляем ID в конец строки
            row.push(sourceRowIndex + 1);

            // Получаем ссылки из колонок L и AJ
            let linkRichTextL = sourceRichTextValues[sourceRowIndex][headers.indexOf('L')];  // Get the RichTextValue from column L
            if (linkRichTextL) {
                let linkUrlL = linkRichTextL.getLinkUrl();
                if (linkUrlL) {
                    row[headers.indexOf('L')] = linkUrlL;
                } else {
                    // Logger.log("Failed to get link from column L");
                }
            }

            let linkRichTextAJ = sourceRichTextValues[sourceRowIndex][headers.indexOf('AJ')];  // Get the RichTextValue from column AJ
            if (linkRichTextAJ) {
                let linkUrlAJ = linkRichTextAJ.getLinkUrl();
                if (linkUrlAJ) {
                    row[headers.indexOf('AJ')] = linkUrlAJ;
                } else {
                    // Logger.log("Failed to get link from column AJ");
                }
            }

            destSheet.appendRow(row);
        }
    }



    // Записываем время генерации данных
    var now = Utilities.formatDate(new Date(), 'Asia/Tbilisi', 'dd/MM/yyyy, HH:mm:ss');
    destSheet.getRange('A1').setValue('Generated at ' + now + ' (Tbilisi, Georgia Timezone)');

    // Logger.log('Process completed');
}

function collectDeveloperCvData() {
    var folderIds = ['1Oia9sf52enMuPCLTV2O5PHvEp72atnWv', '1j39Yb0UEr5VygDmTzbILZ0Sx4Nm8Xs3H'];
    var destSheetName = 'DeveloperCvData';

    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var destSheet = activeSpreadsheet.getSheetByName(destSheetName);
    if (!destSheet) {
        // Если лист не существует, создаем новый
        destSheet = activeSpreadsheet.insertSheet(destSheetName);
    }

    // Очищаем лист
    destSheet.clear();

    // Записываем заголовки
    var headers = ['Имя папки', 'ID папки', 'Имя файла', 'ID файла', 'Дата последнего изменения файла', 'CV link'];
    destSheet.getRange('A2:' + columnToLetter(headers.length) + '2').setValues([headers]);

    // Logger.log("Headers written to destination sheet");

    folderIds.forEach(function(folderId) {
        var folder = DriveApp.getFolderById(folderId);
        var subFolders = folder.getFolders();
        while (subFolders.hasNext()) {
            var subFolder = subFolders.next();
            var folderName = subFolder.getName();
            // Logger.log('Processing folder: ' + folderName);

            // Проверяем, что имя папки соответствует русскому или английскому имени разработчика
            // if (englishNames.includes(folderName) || russianNames.includes(folderName)) {
            var files = subFolder.getFiles();
            while (files.hasNext()) {
                var file = files.next();
                // Logger.log('Processing file: ' + file.getName());

                var row = [
                    subFolder.getName(),
                    subFolder.getId(),
                    file.getName(),
                    file.getId(),
                    file.getLastUpdated(),
                    'https://drive.google.com/file/d/' + file.getId() + '/edit'  // CV link
                ];

                destSheet.appendRow(row);
            }
            // } else {
            //     // Logger.log('Folder name does not match any developer names: ' + folderName);
            // }
        }
    });

    // Записываем время генерации данных
    var now = Utilities.formatDate(new Date(), 'Asia/Tbilisi', 'dd/MM/yyyy, HH:mm:ss');
    destSheet.getRange('A1').setValue('Generated at ' + now + ' (Tbilisi, Georgia Timezone)');

    // Logger.log('Process completed');
}



function collectDeveloperUpworkData() {
    var sourceDocId = '1niuMkbPAOZJ7OCdzIVTwdor6Ug9ljhq2K7_K9ecci9I';
    var sourceSheetName = 'All accounts';
    var destSheetName = 'DeveloperUpworkData';

    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var destSheet = activeSpreadsheet.getSheetByName(destSheetName);
    if (!destSheet) {
        // Если лист не существует, создаем новый
        destSheet = activeSpreadsheet.insertSheet(destSheetName);
    }

    // Очищаем лист
    destSheet.clear();

    var sourceSheet = SpreadsheetApp.openById(sourceDocId).getSheetByName(sourceSheetName);
    var sourceData = sourceSheet.getDataRange().getValues();

    // Записываем заголовки
    var headers = sourceData[0];
    //headers.push('Имя агентства', 'upwork-account');
    destSheet.getRange('A2:' + columnToLetter(headers.length) + '2').setValues([headers]);

    // Logger.log("Headers written to destination sheet");

    // var agencyName = '';
    for(var i = 1; i < sourceData.length; i++) {
        var row = sourceData[i];

        //   // Проверка, является ли строка строкой агентства
        //   if(row.slice(1).every(cell => !cell)) {
        //     agencyName = row[0]; // Обновляем имя агентства
        //   } else {
        //     var upworkAccountMatch = row[0].match(/\(([^)]+)\)/); // Ищем upwork-account в скобках
        //     var upworkAccount = upworkAccountMatch ? upworkAccountMatch[1] : '';
        //     row.push(agencyName, upworkAccount); // Добавляем имя агентства и upwork-account в строку
        destSheet.appendRow(row);
        //   }
    }

    // Записываем время генерации данных
    var now = Utilities.formatDate(new Date(), 'Asia/Tbilisi', 'dd/MM/yyyy, HH:mm:ss');
    destSheet.getRange('A1').setValue('Generated at ' + now + ' (Tbilisi, Georgia Timezone)');

    // Logger.log('Process completed');
}



// Функция для преобразования номера колонки в буквенный эквивалент
function columnToLetter(column) {
    var temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

function getNameTranslations() {
    const developersSheetId = "1VW615PcoaR90HLDD-JQeDmeAcz6DH1T_gCuN17v9C1I";
    const developersSheet = SpreadsheetApp.openById(developersSheetId);
    const nameTranslationsSheet = developersSheet.getSheetByName("Developers english vs russian names");

    const data = nameTranslationsSheet.getDataRange().getValues();
    const nameTranslations = {};

    data.forEach(row => {
        // Предполагается, что английские имена находятся в первом столбце (индекс 0), а русские - во втором (индекс 1)
        const englishName = row[0];
        const russianName = row[1];

        nameTranslations[englishName] = russianName;
    });

    return nameTranslations;
}

function getRusNameTranslations() {
    const developersSheetId = "1VW615PcoaR90HLDD-JQeDmeAcz6DH1T_gCuN17v9C1I";
    const developersSheet = SpreadsheetApp.openById(developersSheetId);
    const nameTranslationsSheet = developersSheet.getSheetByName("Developers english vs russian names");

    const data = nameTranslationsSheet.getDataRange().getValues();
    const nameTranslations = {};

    data.forEach(row => {
        // Предполагается, что английские имена находятся в первом столбце (индекс 0), а русские - во втором (индекс 1)
        const englishName = row[0];
        const russianName = row[1];

        nameTranslations[russianName] = englishName;
    });

    return nameTranslations;
}

function getAllDevelopersCompetenceData() {
    const docId = '1VW615PcoaR90HLDD-JQeDmeAcz6DH1T_gCuN17v9C1I'; // ID вашего документа
    const ss = SpreadsheetApp.openById(docId);

    // Определение листов
    const namesSheet = ss.getSheetByName('Developers english vs russian names');
    const softSkillsSheet = ss.getSheetByName('DeveloperSoftSkillsData');
    const languageSheet = ss.getSheetByName('DeveloperLanguageData');
    const skillsSheet = ss.getSheetByName('DeveloperSkillsData');
    const stackSheet = ss.getSheetByName('DeveloperStackData');

    // Получение данных о соответствии имен
    const namesData = namesSheet.getRange(2, 1, namesSheet.getLastRow() - 1, 2).getValues();
    const nameTranslations = Object.fromEntries(namesData.map(row => [row[0], row[1]]));

    let allDevelopersCompetenceData = {};

    // Обработка SoftSkills
    const softSkillsData = softSkillsSheet.getDataRange().getValues();
    softSkillsData.slice(1).forEach(row => {
        if (row[1] !== "internal" || row[2] !== "оценка") return; // Пропускаем строки

        const englishName = row[0];
        const russianName = nameTranslations[englishName];
        if (!allDevelopersCompetenceData[russianName]) {
            allDevelopersCompetenceData[russianName] = {};
        }

        allDevelopersCompetenceData[russianName]['Обучаемость'] = row[3];
        allDevelopersCompetenceData[russianName]['Стрессоустойчивость'] = row[4];
        allDevelopersCompetenceData[russianName]['Работа в команде'] = row[5];
        allDevelopersCompetenceData[russianName]['Работа с клиентом (командой клиента)'] = row[6];
        allDevelopersCompetenceData[russianName]['Навыки самопрезентации'] = row[7];
        allDevelopersCompetenceData[russianName]['Гибкость мышления'] = row[8];
    });

    // Обработка Language
    const languageData = languageSheet.getDataRange().getValues();
    languageData.slice(1).forEach(row => {
        if (row[1] !== "English") return; // Пропускаем строки, где вторая колонка не "English"
        const englishName = row[0];
        const russianName = nameTranslations[englishName];
        allDevelopersCompetenceData[russianName]['Английский'] = row[2]; // C column for level
    });

    // Обработка Skills и Stacks
    // Это более сложная часть, так как данные могут быть разбросаны по нескольким строкам для одного разработчика
    const skillsData = skillsSheet.getDataRange().getValues();
    const stackData = stackSheet.getDataRange().getValues();

    skillsData.slice(1).forEach(row => {
        const englishName = row[0];
        const russianName = nameTranslations[englishName];
        const skill = row[3] + ' (' + row[4] + ')'; // D column for skill
        if (allDevelopersCompetenceData[russianName]['Инструменты\nБиблиотеки\nСитстемы']) {
            allDevelopersCompetenceData[russianName]['Инструменты\nБиблиотеки\nСитстемы'] += '\n' + skill;
        } else {
            allDevelopersCompetenceData[russianName]['Инструменты\nБиблиотеки\nСитстемы'] = skill;
        }
    });

    stackData.slice(1).forEach(row => {
        const englishName = row[0];
        const russianName = nameTranslations[englishName];
        if (!allDevelopersCompetenceData[russianName]['Stack']) {
            allDevelopersCompetenceData[russianName]['Stack'] = {};
        }
        // F column determines if it's Primary or Additional stack
        const stackType = row[5] === 'Основной' ? 'Основной стек' : 'Дополнительный стек';
        const technologyWithLevel = `${row[2]} (${row[3]})`; // C column for technology, D for level
        if (!allDevelopersCompetenceData[russianName][stackType]) {
            allDevelopersCompetenceData[russianName][stackType] = [];
        }
        allDevelopersCompetenceData[russianName][stackType].push(technologyWithLevel);
    });

    return allDevelopersCompetenceData;
}

function getAllDevelopersCvDataFromSheet() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeveloperCvData');
    var data = sheet.getDataRange().getValues();
    var developersCvData = {};

    data.forEach(function(row, index) {
        if (index === 0) {
            return; // Skip headers
        }

        var developerName = row[0]; // Developer name
        var folderId = row[1]; // Folder ID
        var cvFileName = row[2]; // CV file name
        var cvFileId = row[3]; // File ID
        var lastUpdate = row[4]; // Last update date
        var cvLink = row[5]; // CV link

        if (!developersCvData[developerName]) {
            developersCvData[developerName] = { folders: {} };
        }

        if (!developersCvData[developerName].folders[folderId]) {
            developersCvData[developerName].folders[folderId] = { cvList: [] };
        }

        developersCvData[developerName].folders[folderId].cvList.push({ fileName: cvFileName, fileId: cvFileId, link: cvLink, lastUpdate: lastUpdate });
    });

    return developersCvData;
}


function getAllDevelopersUpworkDataFromSheet() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeveloperUpworkData');
    var data = sheet.getDataRange().getValues();
    var developersUpworkData = {};

    // Получаем данные о соответствии имен с листа "Developers english vs russian names"
    const developersSheetId = "1VW615PcoaR90HLDD-JQeDmeAcz6DH1T_gCuN17v9C1I";
    const developersSheet = SpreadsheetApp.openById(developersSheetId);
    const namesSheet = developersSheet.getSheetByName("Developers english vs russian names");
    var namesData = namesSheet.getDataRange().getValues();
    var namesMap = {};
    namesData.forEach(function(row) {
        namesMap[row[0]] = row[1]; // Сопоставление английского имени и русского имени
    });

    data.forEach(function(row, index) {
        if (index === 0) {
            return; // Skip headers
        }

        var developerName = row[0]; // Имя разработчика на английском
        // Очистить имя разработчика от данных в скобках и пробелов по краям
        developerName = developerName.replace(/\(.*\)/, "").trim();
        // Поменять местами имя и фамилию
        developerName = developerName.split(' ').reverse().join(' ');

        // Поиск соответствующего русского имени
        var russianName = null;
        for (var name in namesMap) {
            if (levenshtein(name, developerName) <= 0.2 * name.length) {
                russianName = namesMap[name];
                break;
            }
        }

        if (russianName) {
            var upworkLink = row[1];
            var upworkAccount = row[0];
            var active = row[2];
            var techStack = row[3];
            var extraTechStack = row[4]

            developersUpworkData[russianName] = {
                upworkLink: upworkLink,
                upworkAccount: upworkAccount,
                active: active,
                techStack: techStack,
                extraTechStack: extraTechStack
            };
        }
    });

    return developersUpworkData;
}



function getAllCandidatesDataFromSheet() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CandidatesData');
    var data = sheet.getDataRange().getValues();
    var candidatesData = {};

    var headers = data[1]; // Получаем заголовки из первой строки

    data.forEach(function(row, index) {
        if (index === 0) {
            return; // Skip headers
        }

        var candidateName = row[0]; // Имя кандидата
        var candidateInfo = []; // Массив для хранения информации о кандидате

        for (var i = 1; i < row.length; i++) {
            candidateInfo.push({ field: headers[i], value: row[i] }); // Добавляем объект с названием поля и значением
        }

        candidatesData[candidateName] = candidateInfo;
    });

    return candidatesData;
}



function generateDeveloperInfoCell(russianName) {
    // Получаем данные разработчика
    var developerData = getDeveloperUpworkData(russianName);

    // Создаем RichTextValueBuilder
    var combinedRichTextBuilder = SpreadsheetApp.newRichTextValue();
    var combinedText = '';

    // Добавляем данные разработчика
    var agency = developerData.agencyName ? 'Agency: ' + developerData.agencyName + '\n' : '';
    var name = developerData.name ? 'Name: ' + developerData.name + '\n' : '';
    var upworkLink = developerData.upworkLink ? 'Upwork Profile: ' + developerData.upworkLink + '\n' : '';
    var busyHours = developerData.busyHours ? 'Busy Hours: ' + developerData.busyHours + '\n' : '';
    var busyProjects = developerData.busyProjects ? 'Busy on Projects: ' + developerData.busyProjects + '\n' : '';
    var RP = developerData.RP ? 'RP: ' + developerData.RP + '\n' : '';
    var rate2021 = developerData.rate2021 ? '2021 Rate: ' + developerData.rate2021 + '\n' : '';
    var remainingData = '';

    for (var key in developerData) {
        if (developerData.hasOwnProperty(key)) {
            // Не выводим уже отображенные данные
            if (['agencyName', 'name', 'upworkLink', 'busyHours', 'busyProjects', 'RP', 'rate2021'].includes(key)) {
                continue;
            }
            remainingData += key.charAt(0).toUpperCase() + key.slice(1) + ': ' + developerData[key] + '\n';
        }
    }

    // Составляем окончательный текст
    combinedText += agency + name + upworkLink + busyHours + busyProjects + RP + rate2021 + remainingData;

    combinedRichTextBuilder.setText(combinedText);  // Устанавливаем текст в RichTextValueBuilder

    // Добавляем ссылку на профиль Upwork, если она доступна
    if (developerData.upworkLink) {
        combinedRichTextBuilder.setLinkUrl(combinedText.indexOf('Upwork Profile:'), combinedText.indexOf('Upwork Profile:') + 'Upwork Profile:'.length, developerData.upworkLink);
    }

    // Построение итогового RichTextValue
    var combinedRichText = combinedRichTextBuilder.build();

    return combinedRichText;
}



function getWeekPlan() {
    var externalSheetId = "1N65NUtqBA855C6K8swmeFQ9HbvIZU4fq4EnhYzvNV7Q"; // Замените на внешний ID таблицы
    var externalSpreadsheet = SpreadsheetApp.openById(externalSheetId);
    var externalSheet = externalSpreadsheet.getSheetByName(sheetName);

    var result = []; // Объект для хранения результатов

    for (var i = 0; i < students.length; i++) {
        var studentName = students[i][0].toString();
        if (studentName == "") continue;

        var trainingHours = 0; // Счетчик часов обучения
        var hrHours = 0; // Счетчик HR часов
        var projectHours = getHoursByNameAndProject(data, studentName); // вызываем функцию, чтобы собрать информацию о проектах и часах

        for (var j = 0; j < 100; j++) {
            for (var k = 0; k < data[j].length; k++) {
                if (data[j][1].toString().startsWith(studentName) && studentName != "") {
                    // Если это часы обучения или HR часы, то добавляем к соответствующему счетчику
                    if (headers[k+2] == "Training") {
                        trainingHours += data[j][k];
                    }
                    else if (headers[k+2] == "HR") {
                        hrHours += data[j][k];
                    }
                }
            }
        }
        // Собираем данные в объект
        result.push({
            name: studentName,
            trainingHours: trainingHours,
            hrHours: hrHours,
            projectHours: projectHours
        });
    }

    // Возвращаем собранный объект
    return result;
}


function updateDeveloperStackData() {
    // Logger.log('Начало обработки файлов');

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('DeveloperStackData');

    // Create or clear the 'DeveloperStackData' sheet
    if (sheet == null) {
        sheet = spreadsheet.insertSheet('DeveloperStackData');
        // Logger.log('Лист "DeveloperStackData" создан');
    } else {
        sheet.clear();
        // Logger.log('Лист "DeveloperStackData" очищен');
    }

    // Fetch data from the existing stack data spreadsheet
    var sourceSpreadsheetId = '1VW615PcoaR90HLDD-JQeDmeAcz6DH1T_gCuN17v9C1I';
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    var sourceSheet = sourceSpreadsheet.getSheetByName('DeveloperStackData');
    var range = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 6); // Adjust column index and count based on actual data layout
    var data = range.getValues();

    // Write data to 'DeveloperStackData'
    sheet.appendRow(['Data gathered at ' + new Date().toLocaleString("en-GB", { timeZone: "Asia/Tbilisi" }) + ' (Tbilisi, Georgia Timezone)']);
    sheet.appendRow(['Developer', 'Тип', 'Технология', 'Уровень', 'Желание/Нежелание', 'Стек']); // Headers
    if (data.length > 0) {
        sheet.getRange(3, 1, data.length, 6).setValues(data);
        // Logger.log('Данные успешно записаны на лист "DeveloperStackData"');
    }

    // Optionally, call a function to correct data if necessary
    correctDataInDeveloperStackData();
}




function getDataFromRange(sheet, range, stackType, developerName) {
    // Logger.log('Сбор данных из диапазона: ' + range);

    var data = sheet.getRange(range).getValues();
    var output = [];

    for (var i = 4; i < data.length; i++) {
        var row = data[i];
        if (row[0] === '') {
            break;
        }
        output.push([developerName, ...row, stackType]);
    }

    // Logger.log('Данные из диапазона ' + range + ' собраны');
    return output;
}

function correctDataInDeveloperStackData() {
    // Logger.log('Начало корректировки данных');

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeveloperStackData');
    var data = sheet.getRange(3, 1, sheet.getLastRow(), 6).getValues();
    var correctedData = [];

    var types = [...new Set(data.map(row => row[1].toLowerCase()))];
    var technologies = [...new Set(data.map(row => row[2].toLowerCase()))];
    //var levels = [...new Set(data.map(row => row[3]))]; // Здесь мы не приводим к нижнему регистру

    // Logger.log('Получение уникальных значений для типов, технологий и уровней');

    var typeCorrections = getCorrections(types, data, 1);
    var technologyCorrections = getCorrections(technologies, data, 2, true, 1); // Установить порог схожести в 1 для технологий
    //var levelCorrections = getCorrections(levels, data, 3, false); // Не применяем toLowerCase для уровней

    // Logger.log('Корректировка данных');

    data.forEach(row => {
        correctedData.push([
            row[0],
            (typeCorrections[row[1].toLowerCase()] || row[1]).trim(), // Применяем .trim()
            (technologyCorrections[row[2].toLowerCase()] || row[2]).trim(), // Применяем .trim()
            row[3].replace(' ', '').trim(), // Удалить пробелы в уровнях, применяем .toLowerCase() и .trim()
            row[4],
            row[5]
        ]);
    });

    sheet.getRange(3, 1, sheet.getLastRow(), 6).setValues(correctedData);

    // Logger.log('Завершено: данные скорректированы и обновлены на листе');
}

// Добавить новый аргумент для определения порога для схожести
function getCorrections(uniqueValues, data, colIndex, lowerCase = true, similarityThreshold = 2) {
    var corrections = {};

    uniqueValues.forEach(value => {
        var similarValues = uniqueValues.filter(val => levenshtein(value, val) <= similarityThreshold);
        var counts = similarValues.map(val => {
            return {
                value: val,
                count: data.reduce((count, row) => count + ((lowerCase ? row[colIndex].toLowerCase() : row[colIndex]) === val ? 1 : 0), 0)
            };
        });
        counts.sort((a, b) => b.count - a.count);
        corrections[value] = counts[0].value;
    });

    return corrections;
}

function gatherKeywords() {
    // Открываем целевой документ по ID
    var targetDocId = "1VW615PcoaR90HLDD-JQeDmeAcz6DH1T_gCuN17v9C1I";
    var targetSs = SpreadsheetApp.openById(targetDocId);

    // Получаем активный документ, куда будем записывать ключевые слова
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Идентификаторы колонок
    var skillsColumn = 'D';
    var stackColumn = 'C';

    // Названия листов
    var skillsSheetName = 'DeveloperSkillsData';
    var stackSheetName = 'DeveloperStackData';

    // Получаем данные из указанных листов и колонок целевого документа
    var skillsData = targetSs.getSheetByName(skillsSheetName).getRange(skillsColumn + "1:" + skillsColumn + targetSs.getSheetByName(skillsSheetName).getLastRow()).getValues();
    var stackData = targetSs.getSheetByName(stackSheetName).getRange(stackColumn + "1:" + stackColumn + targetSs.getSheetByName(stackSheetName).getLastRow()).getValues();

    // Объединяем данные
    var allValues = skillsData.concat(stackData);

    // Преобразуем двумерный массив в одномерный, приводим значения к нижнему регистру и фильтруем пустые значения
    var flattenedValues = allValues.flat().map(function(item) {
        return item.toString().toLowerCase();
    }).filter(function(item) {
        return item !== "";
    });

    // Получаем уникальные значения и сортируем
    var uniqueValues = Array.from(new Set(flattenedValues)).sort();

    // Создаем лист "Keywords", если его еще нет, или очищаем существующий в активном документе
    var keywordSheet = ss.getSheetByName('Keywords');
    if (!keywordSheet) {
        keywordSheet = ss.insertSheet('Keywords');
    } else {
        keywordSheet.clear();
    }

    // Записываем уникальные значения в лист "Keywords" активного документа
    uniqueValues.forEach(function(value, index) {
        keywordSheet.getRange(index + 1, 1).setValue(value);
    });
}


function levenshtein(a, b) {
    if (a.length == 0) return b.length;
    if (b.length == 0) return a.length;

    var matrix = [];

    var i;
    for (i = 0; i <= b.length; i++) {
        matrix[i] = [i];
    }

    var j;
    for (j = 0; j <= a.length; j++) {
        matrix[0][j] = j;
    }

    for (i = 1; i <= b.length; i++) {
        for (j = 1; j <= a.length; j++) {
            if (b.charAt(i-1) == a.charAt(j-1)) {
                matrix[i][j] = matrix[i-1][j-1];
            } else {
                matrix[i][j] = Math.min(matrix[i-1][j-1] + 1, Math.min(matrix[i][j-1] + 1, matrix[i-1][j] + 1));
            }
        }
    }

    return matrix[b.length][a.length];
};

function findClosestMatch(target, array) {
    var minDist = Infinity;
    var match = '';

    for (var i = 0; i < array.length; i++) {
        var dist = levenshtein(target, array[i]);

        if (dist < minDist) {
            minDist = dist;
            match = array[i];
        }
    }

    return match;
};

function insertSumFormulas(all, isLastWeek = false, isBench = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const benchSheetId = "1fpe40DxU-diKV_MfQayPIsBlTGDPBWeBrifCyUGdhy4";
    const benchSpreadsheet = SpreadsheetApp.openById(benchSheetId);

    let reportName = "SALES report"

    if(all)
        reportName = 'ALL report';
    if(isBench)
        reportName = 'SharpDev Bench Report';

    // const ss = SpreadsheetApp.getActiveSpreadsheet();

    // let sheetName = "SALES report"
    // if(all) {
    //   sheetName = 'ALL report';
    // }

    // if(isLastWeek) {
    //   sheetName += ' last week';
    // }

    if(!isBench) sheet = ss.getSheetByName(reportName);
    else sheet = benchSpreadsheet.getSheetByName(reportName);

    var lastRow = sheet.getLastRow();
    var startRow = 7; // Начальная строка, с которой нужно вставлять формулы
    var startColumn = 1; // Начальный столбец, для которого нужно вставить формулы
    var endColumn = 1; // Последний столбец, для которого нужно вставить формулы

    for (var row = startRow; row <= lastRow; row++) {
        for (var column = startColumn; column <= endColumn; column++) {
            var cell = sheet.getRange(row, column);
            var formula = "=SUM(E" + row + ":F" + row + ")";
            cell.setFormula(formula).setVerticalAlignment("middle");
        }
    }
}

function transliterate(name) {
    const ru = ('абвгдеёжзийклмнопрстуфхцчшщъыьэюя').split('');
    const en = ('a,b,v,g,d,e,yo,zh,z,i,y,k,l,m,n,o,p,r,s,t,u,f,h,c,ch,sh,shch,``,y,`,e,yu,ya').split(',');

    function translitWord(word) {
        let newWord = '';
        for(let i = 0; i < word.length; i++) {
            let str = word[i].toLowerCase();
            let index = ru.indexOf(str);
            if(index > -1) {
                newWord += word[i] === word[i].toUpperCase() ? en[index].toUpperCase() : en[index];
            } else {
                newWord += word[i];
            }
        }
        return newWord;
    }

    let [lastName, firstName] = name.split(' ');
    let translitLastName = translitWord(lastName);
    let translitFirstName = translitWord(firstName);

    return `${translitFirstName} ${translitLastName.charAt(0)}.`;
}


function testGetAllDevelopersUpworkDataFromSheet() {
    var developersUpworkData = getAllDevelopersUpworkDataFromSheet();
    // Logger.log(JSON.stringify(developersUpworkData, null, 2));
}

function testDates() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let isLastWeek = false;
    let mondayDate, sundayDate;

    if (isLastWeek) {
        mondayDate = getLastMonday();
        sundayDate = getLastSunday();
    } else {
        mondayDate = getCurrentMonday();
        sundayDate = getCurrentSunday();
    }

    const mondayString = Utilities.formatDate(mondayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();
    const sundayString = Utilities.formatDate(sundayDate, ss.getSpreadsheetTimeZone(), 'dd MMM').toLowerCase();

    let workloadSheetName = mondayDate.getMonth() === sundayDate.getMonth() ?
        `${mondayString.split(" ")[0]}-${sundayString.split(" ")[0]} ${sundayString.split(" ")[1]}` :
        `${mondayString}-${sundayString}`;

    let reportSheetName = mondayDate.getMonth() === sundayDate.getMonth() ?
        `${mondayString.split(" ")[0]}-${sundayString.split(" ")[0]} ${sundayString.split(" ")[1]}` :
        `${mondayString}-${sundayString}`;

    // Logger.log(workloadSheetName + " " + mondayString);
}

// Функция для поиска разработчика на листе DeveloperVacation
function findDeveloperVacation(values, name) {
    for (var i = 2; i <= values.length; i++) {
        if (values[i - 1][0].startsWith(name)) {
            return values[i - 1];
        }
    }

    return -1; // Разработчик не найден
}

// Функция для поиска разработчика на листе DeveloperRate
function findDeveloperRate(values, name) {
    for (var i = 2; i <= values.length; i++) {
        if (values[i - 1][0].startsWith(name)) {
            return values[i - 1];
        }
    }

    return -1; // Разработчик не найден
}

// Функция для поиска разработчика на листе DeveloperProfile (имя на английском)
function findDeveloperProfileLink(values, name) {
    for (var i = 2; i <= values.length; i++) {
        if (values[i - 1][0].startsWith(name)) {
            return values[i - 1][1];
        }
    }

    return -1; // Разработчик не найден

}

function findEnglishName(values, name) {
    // first column is english, second is russian
    for (var i = 2; i <= values.length; i++) {
        if (values[i - 1][1].startsWith(name)) {
            return values[i - 1][0];
        }
    }

    return -1; // Разработчик не найден
}


function cleanOldFiles() {
    var folderId = '15lR-TFQyzeQ7WiZRZhZeeNTQko31nNvV';
    var folder = DriveApp.getFolderById(folderId);
    var now = new Date().getTime();
    var files = folder.getFiles();

    while (files.hasNext()) {
        var file = files.next();
        var dateCreated = file.getDateCreated().getTime();

        if (now - dateCreated > 3600000) { // 3600000 миллисекунд = 1 час
            file.setTrashed(true);
        }
    }
}

function extractIdFromUrl(url) {
    var pattern = /\/d\/([a-zA-Z0-9-_]+)/;
    var matches = url.match(pattern);
    if (matches) {
        return matches[1]; // Возвращает ID документа
    }
    return null; // Возвращает null, если ID не найден
}


// Функция для нахождения моды в массиве
function findMode(array) {
    let frequency = {}; // Объект для хранения частоты каждого значения
    let maxFrequency = 0; // Максимальная частота
    let modes = [];

    for (let item of array) {
        if (frequency[item]) {
            frequency[item]++;
        } else {
            frequency[item] = 1;
        }

        if (frequency[item] > maxFrequency) {
            maxFrequency = frequency[item];
            modes = [item];
        } else if (frequency[item] === maxFrequency) {
            modes.push(item);
        }
    }

    return modes;
}


function collectRecentModifiedFiles() {
    var folderId = '1nS4NWZ51p9OOi_E7Qw4UsvQ5rUdeFD4N';
    var targetSpreadsheetId = '189YZ_AKtBhVBADGksYIjKQCg8h_ky6Bh5tjEzxUWeXY'; // ID целевого документа
    var destSheetName = 'CandidatesRecentModifiedFilesData';

    // Открываем целевой документ по ID
    var activeSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
    var destSheet = activeSpreadsheet.getSheetByName(destSheetName);
    if (!destSheet) {
        destSheet = activeSpreadsheet.insertSheet(destSheetName);
    }

    // Очищаем лист
    destSheet.clear();

    // Записываем заголовки
    var headers = ['Имя файла', 'ID файла', 'Дата последнего изменения файла', 'Ссылка на файл'];
    destSheet.getRange('A1:' + columnToLetter(headers.length) + '1').setValues([headers]);

    // Получаем файлы из папки
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var fileInfo = [];

    while (files.hasNext()) {
        var file = files.next();
        fileInfo.push({
            name: file.getName(),
            id: file.getId(),
            lastUpdated: file.getLastUpdated(),
            link: 'https://drive.google.com/file/d/' + file.getId() + '/view'
        });
    }

    // Сортируем файлы по дате последнего изменения
    fileInfo.sort(function(a, b) {
        return b.lastUpdated - a.lastUpdated; // Сортировка по убыванию
    });

    // Ограничиваем список до первых 50 файлов
    fileInfo = fileInfo.slice(0, 50);

    // Записываем данные в лист
    fileInfo.forEach(function(file, index) {
        var row = [file.name, file.id, file.lastUpdated, file.link];
        destSheet.appendRow(row);
    });

    // Записываем время генерации данных
    var now = Utilities.formatDate(new Date(), 'GMT+4', 'dd/MM/yyyy, HH:mm:ss');
    destSheet.getRange('A' + (fileInfo.length + 2)).setValue('Generated at ' + now + ' (GMT+4 Timezone)');

    // Logger.log('Process completed. Data collected for the most recent 50 modified files.');
}


function columnToLetter(column) {
    // Преобразуем номер колонки в букву
    var temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}


function addNoHoursNotification(date, developer) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = 'NoHoursNotifications';
    var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

    const timeZone = spreadsheet.getSpreadsheetTimeZone();

    var headers = ['Date', 'Developer', 'Status', 'CurrentDateTime', 'Notification', 'UpdatedDateTime'];
    var currentDateTime = new Date();
    var currentHourTbilisi = currentDateTime.getUTCHours() + 4; // UTC + 4 для Тбилиси

    // Если это не рабочий день или текущее время не в диапазоне 11:00-19:00 по Тбилиси, не добавляем уведомление
    if (currentHourTbilisi < 11 || currentHourTbilisi >= 22) {
        Logger.log('Уведомления формируются только в рабочие дни и только с 11:00 до 19:00 по Тбилиси');
        return;
    }

    // Форматируем дату для сравнения
    var formattedDate = formatDate(new Date(date));

    // Если лист новый, добавляем заголовки
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(headers);
    }

    // Получаем все данные из листа, пропуская заголовки
    var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    var data = dataRange.getValues();

    // Проверяем существование записи с данной парой "Дата-Разработчик"
    var exists = false;
    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var storedDate = formatDate(new Date(row[0]));
        var storedDeveloper = row[1].trim();
        if (storedDate === formattedDate && storedDeveloper.toLowerCase() === developer.toLowerCase()) {
            exists = true;
            // Обновляем текущее время в колонке 'UpdatedDateTime'
            sheet.getRange(i + 2, 6).setValue(currentDateTime); // Обновляем конкретную ячейку (колонка 6 - UpdatedDateTime)
            Logger.log('Запись обновлена для ' + developer + ' на ' + date);
            return;
        }
    }

    // Определение статуса
    var status = 'new';

    // Получаем имя разработчика (второе слово после пробела)
    var developerNameParts = developer.trim().split(' ');
    var developerFirstName = developerNameParts[1] || developerNameParts[0]; // Берем первое слово, если нет второго

    // Формирование текста уведомления
    var notificationText = `Привет ${developerFirstName} \nпожалуйста залогируй время за ${Utilities.formatDate(new Date(date), timeZone, 'dd/MM')} по ссылке\nhttps://sharpdev.atlassian.net/plugins/servlet/ac/com.kaanha.reports/jira-reports-timesheet-entry \nВ случае болезни просим сообщить своему руководителю проекта, а также Екатерине Хомутовой / Андрею Ларцеву.\nХорошего дня!`;

    // Добавление записи в лист
    var newRow = [date, developer, status, currentDateTime, notificationText, currentDateTime];
    sheet.appendRow(newRow);

    Logger.log('Добавлено уведомление для ' + developer + ' на ' + date);
}



// Обновляем или добавляем новую запись в таблице уведомлений
function updateOrAddPaidHoursNotification(sheet, date, status, currentTime, notificationText) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timeZone = ss.getSpreadsheetTimeZone();
    const formattedDate = Utilities.formatDate(date, timeZone, 'dd/MM/yyyy');

    const range = sheet.getRange(1, 1, sheet.getLastRow(), 2);
    const values = range.getValues();

    let found = false;

    for (let i = 0; i < values.length; i++) {
        const cellDate = values[i][0];
        const cellStatus = values[i][1];
        const formattedCellDate = Utilities.formatDate(new Date(cellDate), timeZone, 'dd/MM/yyyy');

        if (formattedCellDate === formattedDate) {
            if (cellStatus === 'sent') {
                // If the status is 'sent', do not update
                found = true;
                console.log('Entry for date', formattedCellDate, 'already sent, not updating');
            } else {
                // Update the existing entry if status is not 'sent'
                sheet.getRange(i + 1, 2).setValue(status);
                sheet.getRange(i + 1, 3).setValue(currentTime);
                sheet.getRange(i + 1, 4).setValue(notificationText);
                found = true;
                console.log('Updated entry for date', formattedCellDate);
            }
            break;
        }
    }

    if (!found) {
        // If no existing entry is found, create a new one
        const lastNotificationRow = sheet.getLastRow() + 1;
        sheet.getRange(lastNotificationRow, 1).setValue(formattedDate); // Date
        sheet.getRange(lastNotificationRow, 2).setValue(status); // Status
        sheet.getRange(lastNotificationRow, 3).setValue(currentTime); // CurrentTime
        sheet.getRange(lastNotificationRow, 4).setValue(notificationText); // Notification
        console.log('Added new entry for date', formattedDate);
    }
}

function testSendSlackNotification() {
    sendSlackNotification('test', 'U06KXR5MP0U');
}


function sendSlackNotification(message, channel = 'U06KXR5MP0U') {
    var service = getOAuthService();
    if (service.hasAccess()) {
        var url = 'https://slack.com/api/chat.postMessage';
        var payload = {
            channel: channel,
            text: message
        };

        var params = {
            method: 'post',
            contentType: 'application/json',
            headers: {
                'Authorization': 'Bearer ' + service.getAccessToken()
            },
            payload: JSON.stringify(payload)
        };

        var response = UrlFetchApp.fetch(url, params);
        var result = JSON.parse(response.getContentText());
        if (!result.ok) {
            Logger.log('Error posting message: ' + result.error);
            return null; // Возвращаем null, если ошибка
        } else {
            return { ts: result.ts, channel: result.channel }; // Возвращаем объект с timestamp и каналом
        }
    } else {
        Logger.log('No OAuth access. Please re-authorize.');
        authorize(); // Initiate the authorization process
        return null; // Возвращаем null, если нет доступа
    }
}



function getAccessToken(clientId, clientSecret, refreshToken) {
    var url = 'https://slack.com/api/oauth.v2.access';
    var payload = {
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token'
    };

    var options = {
        method: 'post',
        contentType: 'application/x-www-form-urlencoded',
        payload: payload
    };

    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    if (result.ok) {
        return result.access_token;
    } else {
        Logger.log('Error getting access token: ' + result.error);
        return null;
    }
}

function processNotifications() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var noHoursNotificationsSheetName = 'NoHoursNotifications';
    var paidHoursNotificationsSheetName = 'PaidHoursNotifications';

    processNoHoursNotifications(spreadsheet, noHoursNotificationsSheetName);
    processPaidHoursNotifications(spreadsheet, paidHoursNotificationsSheetName);
}

function processNoHoursNotifications(spreadsheet, sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    var slackSheet = SpreadsheetApp.openById('1VW615PcoaR90HLDD-JQeDmeAcz6DH1T_gCuN17v9C1I').getSheetByName('Developers english vs russian names');
    var defaultChannel = 'U06KXR5MP0U'; // Andrey Lartsev personal

    if (!sheet || !slackSheet) {
        Logger.log('Не найден лист ' + sheetName + ' или Developers english vs russian names');
        return;
    }

    var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    var data = dataRange.getValues();

    var slackData = slackSheet.getDataRange().getValues();
    var slackUsers = {};
    slackData.forEach(function(row) {
        slackUsers[row[1]] = row[2]; // Assuming Russian name is in the second column and Slack account is in the third column
    });

    var now = new Date();
    var tbilisiTimezone = 'Asia/Tbilisi';
    var currentHour = Utilities.formatDate(now, tbilisiTimezone, 'HH');

    if (currentHour >= 11 && currentHour < 22) {
        data.forEach(function(row, index) {
            var date = row[0];
            var developer = row[1];
            var status = row[2];
            var currentDateTime = row[3];
            var notification = row[4];

            if (status === 'new') {
                var channel = defaultChannel;
                var userId = slackUsers[developer];
                if (userId) {
                    channel = userId;
                }

                var message = notification;
                sendSlackNotification(message, channel);
                sendSlackNotification('Сообщение для ' + developer + '\n' + message, defaultChannel); // Дублируем уведомление на канал Андрея Лартсева

                // Обновляем статус на "sent"
                sheet.getRange(index + 2, 3).setValue('sent');
                sheet.getRange(index + 2, 3).setNote('Отправлено в Slack: ' + new Date());
            }
        });
    } else {
        Logger.log('Текущее время не входит в интервал 11:00-22:00 по времени Тбилиси.');
    }
}


function processPaidHoursNotifications(spreadsheet, sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    var channel = 'C075BGFNNKV'; // notify-paid-hours-daily

    if (!sheet) {
        Logger.log('Лист ' + sheetName + ' не найден');
        return;
    }

    var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    var data = dataRange.getValues();

    var now = new Date();
    var tbilisiTimezone = 'Asia/Tbilisi';
    var currentHour = Utilities.formatDate(now, tbilisiTimezone, 'HH');

    if (currentHour >= 12 && currentHour < 22) {
        data.forEach(function(row, index) {
            var date = row[0];
            var status = row[1];
            var currentDateTime = row[2];
            var notification = row[3];

            if (status === 'new') {
                var message = notification;
                sendSlackNotification(message,channel);

                // Обновляем статус на "sent"
                sheet.getRange(index + 2, 2).setValue('sent');
                sheet.getRange(index + 2, 2).setNote('Отправлено в Slack: ' + new Date());
            }
        });
    } else {
        Logger.log('Текущее время не входит в интервал 12:00-22:00 по времени Тбилиси.');
    }
}

function updateSlackUsersSheet() {
    var service = getOAuthService();
    if (!service.hasAccess()) {
        Logger.log('No OAuth access. Please re-authorize.');
        return;
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SlackUsers');
    if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('SlackUsers');
    }
    sheet.clear();  // Очистим лист перед обновлением данных
    sheet.getRange(1, 1, 1, 2).setValues([['ID', 'Name']]); // Заголовки столбцов

    var nextCursor = '';
    var users = [];
    var attempts = 0;

    do {
        Logger.log('Fetching Slack users with cursor: ' + nextCursor);
        var response = fetchSlackUsers(service, nextCursor);
        if (response.ok) {
            Logger.log('Fetched ' + response.members.length + ' users');
            response.members.forEach(function(member) {
                if (member.profile) {
                    users.push([member.id, member.profile.real_name]);
                } else {
                    Logger.log('Skipped user without profile: ' + member.id + ' (' + member.name + ')');
                }
            });

            nextCursor = response.response_metadata.next_cursor;
            Logger.log('Next cursor: ' + nextCursor);
            Utilities.sleep(1000); // Задержка между запросами для предотвращения rate limit
        } else {
            Logger.log('Error fetching user list: ' + response.error);
            if (response.error === 'ratelimited') {
                Logger.log('Rate limit exceeded. Retrying in 1 minute.');
                Utilities.sleep(60000); // Задержка на 1 минуту
            } else {
                break;
            }
        }
        attempts++;
    } while (nextCursor && attempts < 10);

    Logger.log('Total users fetched: ' + users.length);

    if (users.length > 0) {
        sheet.getRange(2, 1, users.length, 2).setValues(users); // Вставляем данные
        Logger.log('Users data written to SlackUsers sheet');
    } else {
        Logger.log('No users data to write to SlackUsers sheet');
    }
}

function checkSlackAccounts() {
    var developersSpreadsheetId = '1VW615PcoaR90HLDD-JQeDmeAcz6DH1T_gCuN17v9C1I';
    var developersSpreadsheet = SpreadsheetApp.openById(developersSpreadsheetId);
    var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Проверка существования листов
    var developersSheet = developersSpreadsheet.getSheetByName('Developers english vs russian names');
    if (!developersSheet) {
        Logger.log("Лист 'Developers english vs russian names' не найден");
        return;
    }

    var slackUsersSheet = currentSpreadsheet.getSheetByName('SlackUsers');
    if (!slackUsersSheet) {
        Logger.log("Лист 'SlackUsers' не найден");
        return;
    }

    // Получаем данные из листов
    var developersData = developersSheet.getDataRange().getValues();
    var slackUsersData = slackUsersSheet.getDataRange().getValues();

    // Создаем словарь для быстрого поиска соответствий имени и аккаунта
    var nameToAccountMap = {};
    for (var i = 1; i < developersData.length; i++) { // Пропускаем заголовок
        var name = developersData[i][0];
        if (name) {
            var nameParts = name.split(' ');
            if (nameParts.length > 1) {
                var reversedName = nameParts[1] + ' ' + nameParts[0];
                nameToAccountMap[reversedName.toLowerCase()] = developersData[i][1];
            }
        }
    }

    // Проверяем соответствие и записываем результат в колонку C
    for (var j = 1; j < slackUsersData.length; j++) { // Пропускаем заголовок
        var slackUser = slackUsersData[j][1]; // Имя в колонке B
        if (slackUser) {
            var slackUserName = slackUser.toLowerCase();
            if (nameToAccountMap[slackUserName]) {
                slackUsersSheet.getRange(j + 1, 3).setValue(nameToAccountMap[slackUserName]); // Записываем соответствие в колонку C
            } else {
                slackUsersSheet.getRange(j + 1, 3).setValue('error'); // Записываем 'error', если соответствие не найдено
            }
        }
    }
}



function fetchSlackUsers(service, cursor) {
    var url = 'https://slack.com/api/users.list?limit=200' + (cursor ? '&cursor=' + cursor : '');
    var params = {
        method: 'get',
        contentType: 'application/json',
        headers: {
            'Authorization': 'Bearer ' + service.getAccessToken()
        },
        muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, params);
    Logger.log('Response code: ' + response.getResponseCode());
    var responseText = response.getContentText();
    Logger.log('Response text length: ' + responseText.length); // Логирование длины текста ответа для избежания превышения размера логов
    return JSON.parse(responseText);
}

// // Рекрутинг отчеты
// function sendDailyAndWeeklyReports() {
//   try {
//     Logger.log("Начало выполнения sendDailyAndWeeklyReports");

//     var spreadsheetId = '189YZ_AKtBhVBADGksYIjKQCg8h_ky6Bh5tjEzxUWeXY';
//     var sheetLastWeek = SpreadsheetApp.openById(spreadsheetId).getSheetByName('LastWeekReport');
//     var sheetCurrentWeek = SpreadsheetApp.openById(spreadsheetId).getSheetByName('CurrentWeekReport');

//     var today = new Date();
//     var dayOfWeek = today.getDay();
//     var yesterday = new Date(today);
//     var lastFriday = new Date(today);
//     yesterday.setDate(today.getDate() - 1);
//     lastFriday.setDate(today.getDate() - (today.getDay() + 2)); // Если понедельник, то lastFriday будет пятницей

//     var formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd.MM.yyyy");
//     var formattedYesterday = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "dd.MM.yyyy");
//     var formattedLastFriday = Utilities.formatDate(lastFriday, Session.getScriptTimeZone(), "dd.MM.yyyy");

//     Logger.log("Сегодня: " + formattedToday);
//     Logger.log("Вчера: " + formattedYesterday);
//     Logger.log("Последняя пятница: " + formattedLastFriday);

//     var dailyReport = generateDailyReport(sheetCurrentWeek, sheetLastWeek, formattedYesterday, formattedToday, dayOfWeek, formattedLastFriday);
//     var weeklyReport = '';

//     if(dayOfWeek > 5) {
//       Logger.log("Сегодня выходной день. Не отправляем уведомлений.");
//       return;
//     }

//     if (dayOfWeek === 1) { // Понедельник
//       weeklyReport = generateWeeklyReport(sheetLastWeek);
//     }

//     if (weeklyReport) {
//       Logger.log("Отправка еженедельного отчета");
//       sendSlackNotification(weeklyReport, "C075519G43Y");//D06KEQM525D    , "C075519G43Y"
//     }

//     if (dailyReport) {
//       Logger.log("Отправка ежедневного отчета");
//       sendSlackNotification(dailyReport, "C075519G43Y"); //D06KEQM525D    , "C075519G43Y"
//     } else {
//       Logger.log("Ежедневный отчет не отправлен, так как нет данных за вчера и сегодня");
//     }

//     Logger.log("Завершение выполнения sendDailyAndWeeklyReports");
//   } catch (error) {
//     Logger.log("Ошибка: " + error.message);
//   }
// }

function sendDailyAndWeeklyReports() {
    try {
        Logger.log("Начало выполнения sendDailyAndWeeklyReports");

        var spreadsheetId = '189YZ_AKtBhVBADGksYIjKQCg8h_ky6Bh5tjEzxUWeXY';
        var sheetLastWeek = SpreadsheetApp.openById(spreadsheetId).getSheetByName('LastWeekReport');
        var sheetCurrentWeek = SpreadsheetApp.openById(spreadsheetId).getSheetByName('CurrentWeekReport');

        var today = new Date();
        var dayOfWeek = today.getDay();
        var yesterday = new Date(today);
        var lastFriday = new Date(today);
        yesterday.setDate(today.getDate() - 1);
        lastFriday.setDate(today.getDate() - (today.getDay() + 2)); // Если понедельник, то lastFriday будет пятницей

        var formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd.MM.yyyy");
        var formattedYesterday = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "dd.MM.yyyy");
        var formattedLastFriday = Utilities.formatDate(lastFriday, Session.getScriptTimeZone(), "dd.MM.yyyy");

        Logger.log("Сегодня: " + formattedToday);
        Logger.log("Вчера: " + formattedYesterday);
        Logger.log("Последняя пятница: " + formattedLastFriday);

        var dailyReport = generateDailyReport(sheetCurrentWeek, sheetLastWeek, formattedYesterday, formattedToday, dayOfWeek, formattedLastFriday);
        var weeklyReport = '';

        if (dayOfWeek > 5) {
            Logger.log("Сегодня выходной день. Не отправляем уведомлений.");
            return;
        }

        if (dayOfWeek === 1) { // Понедельник
            weeklyReport = generateWeeklyReport(sheetLastWeek);
        }

        if (dailyReport) {
            Logger.log("Отправка ежедневного отчета");
            sendOrUpdateSlackMessage(dailyReport, 'recruit-daily','C075519G43Y', formattedToday); // Используем функцию sendOrUpdateSlackMessage U06KXR5MP0U
        } else {
            Logger.log("Ежедневный отчет не отправлен, так как нет данных за вчера и сегодня");
        }

        if (weeklyReport) {
            Logger.log("Отправка еженедельного отчета");
            sendOrUpdateSlackMessage(weeklyReport, 'recruit-weekly','C075519G43Y', formattedToday); // Используем функцию sendOrUpdateSlackMessage U06KXR5MP0U
        }

        Logger.log("Завершение выполнения sendDailyAndWeeklyReports");
    } catch (error) {
        Logger.log("Ошибка: " + error.message);
    }
}

function isReportAlreadySent(reportType, date) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SlackMessages') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('SlackMessages');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === reportType && data[i][4] === date) {
            return true;
        }
    }
    return false;
}

//'U06KXR5MP0U'
function sendOrUpdateSlackMessage(text, reportType, primaryChannel = 'U06KXR5MP0U', formattedToday) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SlackMessages') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('SlackMessages');

    // Проверка существования заголовков
    if (sheet.getLastRow() === 0) {
        sheet.appendRow(['ReportType', 'Timestamp', 'PrimaryChannel', 'SecondaryChannel', 'Message', 'Date']);
    }

    // Поиск timestamp и secondaryChannel для отчета
    var data = sheet.getDataRange().getValues();
    var timestamp = '';
    var secondaryChannel = '';
    var rowIndex = -1;
    Logger.log("Начало поиска timestamp для отчета: " + reportType + " на канале: " + primaryChannel + " с датой: " + formattedToday);
    for (var i = 1; i < data.length; i++) {
        Logger.log("Проверка строки " + i + ": " + JSON.stringify(data[i]));
        if (data[i][0] === reportType) {
            Logger.log("Тип отчета совпадает для строки " + i);
        }
        if (data[i][2] === primaryChannel) {
            Logger.log("Primary канал совпадает для строки " + i);
        } else {
            Logger.log("Primary канал не совпадает: ожидалось " + primaryChannel + ", а получено " + data[i][2]);
        }
        if (data[i][5] === formattedToday) {
            Logger.log("Дата совпадает для строки " + i);
        }
        if (data[i][0] === reportType && data[i][2] === primaryChannel && data[i][5] === formattedToday) {
            timestamp = data[i][1];
            secondaryChannel = data[i][3];
            rowIndex = i;
            Logger.log("Найдено совпадение в строке " + i + ": timestamp = " + timestamp + ", secondaryChannel = " + secondaryChannel);
            break;
        }
    }
    if (timestamp === '') {
        Logger.log("Не найдено совпадение для отчета: " + reportType + " на канале: " + primaryChannel + " с датой: " + formattedToday);
    }

    Logger.log('Текущий timestamp для отчета ' + reportType + ': ' + timestamp);

    var service = getOAuthService();
    if (service.hasAccess()) {
        if (timestamp) {
            // Проверяем, изменился ли текст сообщения
            if (data[rowIndex][4] === text) {
                Logger.log('Текст сообщения не изменился, обновление не требуется.');
                return;
            }

            // Обновление существующего сообщения
            Logger.log('Обновление сообщения в канале ' + secondaryChannel + ' с timestamp: ' + timestamp);
            var response = UrlFetchApp.fetch('https://slack.com/api/chat.update', {
                method: 'post',
                contentType: 'application/json',
                headers: {
                    'Authorization': 'Bearer ' + service.getAccessToken(),
                },
                payload: JSON.stringify({
                    channel: secondaryChannel,
                    ts: timestamp,
                    text: text
                })
            });
            var result = JSON.parse(response.getContentText());
            if (!result.ok) {
                Logger.log('Ошибка обновления сообщения: ' + result.error);
                if (result.error === 'message_not_found') {
                    Logger.log('Сообщение не найдено.');
                    // var ts = sendSlackNotification(text, primaryChannel);
                    // if (ts && ts.ts) {
                    //     sheet.appendRow([reportType, ts.ts.toString(), primaryChannel, ts.channel, text, formattedToday]);
                    //     Logger.log('Добавлена новая строка: ' + [reportType, ts.ts.toString(), primaryChannel, ts.channel, text, formattedToday].join(', '));
                    // }
                }
            } else {
                Logger.log('Сообщение успешно обновлено, timestamp: ' + timestamp);
                sheet.getRange(rowIndex + 1, 4).setValue(text); // Обновляем текст в таблице
                sheet.getRange(rowIndex + 1, 5).setValue(formattedToday); // Обновляем дату в таблице
            }
        } else {
            // Отправка нового сообщения
            Logger.log('Отправка нового сообщения в канал ' + primaryChannel);
            var ts = sendSlackNotification(text, primaryChannel);
            if (ts && ts.ts) {
                sheet.appendRow([reportType, ts.ts.toString(), primaryChannel, ts.channel, text, formattedToday]);
                Logger.log('Добавлена новая строка: ' + [reportType, ts.ts.toString(), primaryChannel, ts.channel, text, formattedToday].join(', '));
            }
        }
    } else {
        Logger.log('Нет доступа OAuth. Пожалуйста, повторно авторизуйтесь.');
        authorize(); // Initiate the authorization process
    }
}


function sendSlackNotification(message, channel = 'U06KXR5MP0U') {
    var service = getOAuthService();
    if (service.hasAccess()) {
        var url = 'https://slack.com/api/chat.postMessage';
        var payload = {
            channel: channel,
            text: message
        };

        var params = {
            method: 'post',
            contentType: 'application/json',
            headers: {
                'Authorization': 'Bearer ' + service.getAccessToken()
            },
            payload: JSON.stringify(payload)
        };

        var response = UrlFetchApp.fetch(url, params);
        var result = JSON.parse(response.getContentText());
        if (!result.ok) {
            Logger.log('Ошибка отправки сообщения: ' + result.error);
            return null; // Возвращаем null, если ошибка
        } else {
            Logger.log('Сообщение отправлено, timestamp: ' + result.ts);
            return { ts: result.ts, channel: result.channel }; // Возвращаем объект с timestamp и каналом
        }
    } else {
        Logger.log('Нет доступа OAuth. Пожалуйста, повторно авторизуйтесь.');
        authorize(); // Initiate the authorization process
        return null; // Возвращаем null, если нет доступа
    }
}


function generateDailyReport(sheetCurrentWeek, sheetLastWeek, yesterday, today, dayOfWeek, lastFriday) {
    try {
        Logger.log("Начало выполнения generateDailyReport");

        var data = sheetCurrentWeek.getDataRange().getValues();
        var dataLastWeek = sheetLastWeek.getDataRange().getValues();
        var dailyReport = "*Рекрутинг Daily " + today + "*\n";
        var yesterdayData = '';
        var todayData = '';


        if (dayOfWeek === 1) { // Понедельник
            Logger.log("Генерация отчета за пятницу");
            yesterdayData = extractReportData(dataLastWeek, lastFriday);
            dailyReport += "*В пятницу (" + lastFriday + ")*\n" + yesterdayData;
        } else {
            Logger.log("Генерация отчета за вчера");
            yesterdayData = extractReportData(data, yesterday, false);
            dailyReport += "*Вчера (" + yesterday + ")*\n" + yesterdayData;
        }

        Logger.log("Генерация отчета за сегодня");
        todayData = extractReportData(data, today, true);
        dailyReport += "*Сегодня (" + today + ")*\n" + todayData;

        if (!yesterdayData && !todayData) {
            Logger.log("Нет данных для отчета за вчера и сегодня");
            return ''; // Если данных нет, возвращаем пустую строку
        }

        Logger.log("Завершение выполнения generateDailyReport");
        return dailyReport;
    } catch (error) {
        Logger.log("Ошибка в generateDailyReport: " + error.message);
        throw error;
    }
}

function generateWeeklyReport(sheet) {
    try {
        Logger.log("Начало выполнения generateWeeklyReport");

        var data = sheet.getDataRange().getValues();
        var weeklyReport = "*Рекрутинг Weekly Report*\n";
        weeklyReport += extractWeeklyReportData(data);

        Logger.log("Завершение выполнения generateWeeklyReport");

        return weeklyReport;
    } catch (error) {
        Logger.log("Ошибка в generateWeeklyReport: " + error.message);
        throw error;
    }
}

function extractWeeklyReportData(data) {
    try {
        Logger.log("Начало выполнения extractWeeklyReportData");

        var reportData = '';

        for (var i = 5; i < data.length; i++) { // Начинаем с 5 строки
            Logger.log("Обработка строки " + (i + 1) + ": " + JSON.stringify(data[i]));


            if (typeof data[i][1] === 'object') {
                var cellDate = Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "dd.MM.yyyy");
                if (cellDate) {
                    reportData += "\n*"+ cellDate + "*\n";
                    reportData += extractReportData(data, data[i][1]);
                }
            }

            if (data[i][1] && typeof data[i][1] === 'string' && data[i][1].indexOf("Проведено") !== -1) {
                Logger.log("Конец отчета найден");
                reportData += "\n*" + data[i].join('') + "*\n";
                for (var j = i + 1; j < data.length; j++) {
                    if (data[j][1]) {
                        reportData += "_" + data[j][1] + "_  " + data[j][2] + "\n";
                    }
                }
                break;
            }
        }

        Logger.log("Завершение выполнения extractWeeklyReportData");
        return reportData;
    } catch (error) {
        Logger.log("Ошибка в extractWeeklyReportData: " + error.message);
        throw error;
    }
}

function extractReportData(data, date, isToday = false) {
    try {
        Logger.log("Начало выполнения extractReportData для даты: " + date);

        var reportData = '';
        var inReportSection = false;
        var dateString = (typeof date === 'object') ? Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.MM.yyyy") : date;
        var hasDataForDate = false;

        for (var i = 5; i < data.length; i++) { // Начинаем с 5 строки
            Logger.log("Обработка строки " + (i + 1) + ": " + JSON.stringify(data[i]));

            if (data[i][1]) {
                if (typeof data[i][1] === 'object') {
                    var cellDate = Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "dd.MM.yyyy");
                    if (!dateString || cellDate === dateString) {
                        inReportSection = true;
                        hasDataForDate = false;
                        Logger.log("Дата совпала (объект): " + cellDate);
                        continue;
                    }
                } else if (typeof data[i][1] === 'string') {
                    Logger.log("Проверка строки на наличие даты: " + data[i][1]);
                    if (data[i][1].indexOf(dateString) !== -1) {
                        inReportSection = true;
                        hasDataForDate = false;
                        Logger.log("Дата совпала (строка): " + data[i][1]);
                        continue;
                    }
                }
            }

            if (inReportSection) {
                if (data[i][1] && data[i][2] && data[i][3] && data[i][4] && data[i][5]) {
                    Logger.log("Обработка данных кандидата: " + data[i][1] + " " + data[i][2] + " " + data[i][3] + " " + data[i][4]);
                    var comments = extractRelevantComments(data[i][5], dateString, isToday);
                    if (comments) {
                        var commentsLines = data[i][5].split('\n');
                        var isDateInSecondLineOrLater = commentsLines.some(function(line, index) {
                            return index > 0 && line.indexOf(dateString) !== -1;
                        });

                        if (isDateInSecondLineOrLater) {
                            reportData += ">_• " + data[i][1] + "\t" + data[i][2] + " " + comments + "_\n";
                        } else {
                            reportData += "• " + data[i][1] + "\t" + data[i][2] + "\t" + data[i][4] + "\t" + comments + "\n";
                        }
                    }
                    hasDataForDate = true;
                } else {
                    Logger.log("Пропуск строки: неполные данные или пустая строка");
                }
            }



            // Окончание секции данных для текущей даты
            if (inReportSection && !data[i][1] && (i + 1 >= data.length || (data[i + 1][1] && typeof data[i + 1][1] === 'string' && data[i + 1][1].match(/\d{2}\.\d{2}\.\d{4}/)))) {
                Logger.log("Конец секции данных для даты: " + date);
                if (!hasDataForDate) {
                    reportData += "Нет собеседований\n";
                }
                inReportSection = false;
            }

            // Прерывание на конце отчета
            if (data[i][1] && typeof data[i][1] === 'string' && data[i][1].indexOf("Проведено") !== -1) {
                Logger.log("Конец отчета найден");
                break;
            }
        }

        Logger.log("Завершение выполнения extractReportData");
        return reportData;
    } catch (error) {
        Logger.log("Ошибка в extractReportData: " + error.message);
        throw error;
    }
}

function extractRelevantComments(comments, date) {
    try {
        Logger.log("Начало выполнения extractRelevantComments для даты: " + date);

        var lines = comments.split('\n');
        var relevantComments = lines.filter(function(line) {
            return line.indexOf(date) !== -1;
        });

        var filteredComments = relevantComments.map(function(comment) {
            var parts = comment.split(' ');
            parts.shift(); // Убираем дату
            parts.shift(); // Убираем 'в'
            return parts.join(' '); // Возвращаем комментарий без даты и времени
        }).filter(function(comment) {
            return comment.trim() !== '';
        });

        Logger.log("Завершение выполнения extractRelevantComments");

        return filteredComments.join('\n');
    } catch (error) {
        Logger.log("Ошибка в extractRelevantComments: " + error.message);
        throw error;
    }
}




