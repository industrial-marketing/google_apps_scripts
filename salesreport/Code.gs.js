var monthNamesShort = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
const nameTranslations = getNameTranslations();


function saveOAuthToken() {
    var token = ScriptApp.getOAuthToken();
    PropertiesService.getScriptProperties().setProperty('OAUTH_TOKEN', token);
}


function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Обновление')
        .addItem('Получить токен', 'saveOAuthToken')
        .addItem('Обновить SALES report', 'generateSalesReportCheckOAuth')
        .addItem('Обновить ALL report', 'generateAllReportCheckOAuth')
        .addItem('Обновить ALL report last week', 'generateAllReportLastWeekCheckOAuth')
        .addItem('Обновить "Keywords"', 'gatherKeywords')
        .addItem('Обновить "Scrum files for current week"', 'gatherDataInSheetCheckOAuth')
        .addItem('Обновить "Scrum files for last week"', 'gatherDataInSheetLastWeekCheckOAuth')
        .addItem('Обновить "Competences"', 'copyDataToCompetencesSheetCheckOAuth')
        .addItem('Обновить "DeveloperStackData"', 'updateDeveloperStackDataCheckOAuth')
        .addToUi();
    ui.createMenu('Фильтры')
        .addItem('Показать всё', 'showAllRows')
        .addItem('Только бенч', 'showOnlyBenchRows')
        .addItem('Выбор стеков', 'showStacksDialog')
        .addItem('Поиск', 'showSearchDialog')
        // .addItem('Сортировать по A-Z', 'sortDataAscending')
        // .addItem('Сортировать по Z-A', 'sortDataDescending')
        .addToUi();
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
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
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
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
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
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
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
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    Logger.log(sheetName);

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

    Logger.log(activeStacks); // Выводим activeStacks в Logger
    return activeStacks;
}


function enableStack(stackName) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();
    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var headerRow = sheet.getRange("5:5");
    var headerValues = headerRow.getValues()[0];
    var columnIndex = headerValues.indexOf(stackName) + 1;

    var filter = sheet.getFilter() || sheet.getRange(6, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).createFilter();
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
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
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
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
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
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    showAllRows();

    var dataRange = sheet.getRange(7, 1, lastRow - 6, lastColumn); // Мы начинаем с 6-й строки и идем до последней строки
    var values = dataRange.getValues();
    sheet.showRows(1, values.length);

    var notesColumnI = sheet.getRange(7, 9, lastRow - 6).getNotes(); // Получаем заметки из колонки K (11)

    if(query) {
        sheet.getRange("H3").setValue("Поиск по запросу:").setBackground('black').setFontColor('white');
        sheet.getRange("I3").setValue(query).setBackground('black').setFontColor('white');
    } else {
        sheet.getRange("H3").setValue("").setBackground('white').setFontColor('black');
        sheet.getRange("I3").setValue("").setBackground('white').setFontColor('black');
    }


    // Скрываем все строки в диапазоне данных
    sheet.hideRows(7, values.length);

    // Преобразуем запрос в массив слов
    var queryWords = query.toLowerCase().split(" ");

    for(var i = 0; i < values.length; i++) {
        // Если строка содержит все слова из запроса, показываем ее
        var rowContent = values[i].join(" ") + " " + notesColumnI[i][0]; // Добавляем содержимое заметки к строке
        var rowContentLower = rowContent.toLowerCase();

        var containsAllWords = queryWords.every(function(word) {
            return rowContentLower.includes(word);
        });

        if(containsAllWords) {
            sheet.showRows(i + 7); // Показываем строки, начиная с 7-й строки
        }
    }
}

function showAllRows() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
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
    range.createFilter();

    // Сбрасываем поиск
    sheet.getRange("H3:J3").clearContent().setBackground('white').setFontColor('black');

    return true;
}


function getCurrentSearchQuery() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var query = sheet.getRange("I3").getValue();
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
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
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
    sheet.getRange("H3").clearContent();
    sheet.getRange("I3").clearContent();
    sheet.getRange("H3").setValue("").setBackground('white').setFontColor('black');
    sheet.getRange("I3").setValue("").setBackground('white').setFontColor('black');
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




function updateDeveloperStackDataCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию updateDeveloperStackData()
        updateDeveloperStackData();
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
        } else {
            // Пользователь нажал Cancel, не выполняем функцию parseData()
            return;
        }
    }
}

function copyDataToCompetencesSheetCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию copyDataToCompetencesSheet()
        copyDataToCompetencesSheet();
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
        } else {
            // Пользователь нажал Cancel, не выполняем функцию parseData()
            return;
        }
    }
}

function generateAllReportCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию generateSalesReport(true)
        generateSalesReport(true);
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
        } else {
            // Пользователь нажал Cancel, не выполняем функцию parseData()
            return;
        }
    }
}


function generateAllReportLastWeekCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию generateAllReportLastWeek(true)
        generateSalesReport(true,true);
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
        } else {
            return;
        }
    }
}


function generateSalesReportCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию generateSalesReport()
        generateSalesReport();
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
        } else {
            // Пользователь нажал Cancel, не выполняем функцию parseData()
            return;
        }
    }
}


function gatherDataInSheetCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию updateWeekPlan()
        gatherDataInSheet();
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
        } else {
            // Пользователь нажал Cancel, не выполняем функцию parseData()
            return;
        }
    }
}

function gatherDataInSheetLastWeekCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию updateWeekPlan()
        gatherDataInSheet(true);
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
        } else {
            // Пользователь нажал Cancel, не выполняем функцию parseData()
            return;
        }
    }
}

function generateAllReportLastWeek() {
    generateSalesReport(true, true);
}

function generateAllReport() {
    generateSalesReport(true);
}

function generateSalesReport(all = false, isLastWeek = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let reportName = "SALES report"
    if(all) {
        reportName = 'ALL report';
    }

    // Если флаг isLastWeek установлен в true, добавляем 'last week' к имени отчета
    if(isLastWeek) {
        reportName += ' last week';
    }

    // Проверяем, запущена ли функция на правильном листе
    if (reportName !== 'ALL report' && reportName !== 'SALES report' && reportName !== 'ALL report last week') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    reportSheet = ss.getSheetByName(reportName);

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

    const mondayString = Utilities.formatDate(mondayDate, ss.getSpreadsheetTimeZone(), 'd MMM').toLowerCase();
    const sundayString = Utilities.formatDate(sundayDate, ss.getSpreadsheetTimeZone(), 'd MMM').toLowerCase();

    let workloadSheetName = mondayDate.getMonth() === sundayDate.getMonth() ?
        `${mondayString.split(" ")[0]}-${sundayString.split(" ")[0]} ${sundayString.split(" ")[1]}` :
        `${mondayString}-${sundayString}`;

    const workloadSheet = workloadSpreadsheet.getSheetByName(workloadSheetName);
    if (!workloadSheet) {
        SpreadsheetApp.getUi().alert(`Cannot find sheet "${workloadSheetName}" in the workload spreadsheet.`);
        return;
    }

    let developers = getDevelopers(workloadSheet, all);

    // Get data for all developers
    let allAllocationData = getAllocationData(developers, isLastWeek);

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
        allAllocationData[developerName].allocationList += " | vacation (" + hours + ")";
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
        var allDevelopersCvData = null;  // Закрытая переменная

        return function(developerName) {
            if (!allDevelopersCvData) {
                allDevelopersCvData = getAllDevelopersCvDataFromSheet();  // Получить данные только при первом вызове
            }

            // Возвращаемые данные для указанного разработчика или пустой объект, если такого разработчика нет
            return allDevelopersCvData[developerName] || { folderId: null, cvList: [], candidateData: {} };
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



    Logger.log(developers.length);


    // Initialize report
    reportSheet.clearContents();
    reportSheet.getRange('B3').setValue( reportName + ` for ${mondayString} - ${sundayString}`).setFontSize(20);
    reportSheet.getRange('K3').setValue('для сортировки выделите колонку и нажмите "Сортировать" или используйте дополнительные инструменты поиска в меню "Фильтры"').setFontSize(9);

    // Initialize the header row
    reportSheet.getRange('B5').setValue('Developer').setVerticalAlignment("middle");
    reportSheet.setColumnWidth(2, 200);
    reportSheet.getRange('C5').setValue('Location').setVerticalAlignment("middle");
    reportSheet.setColumnWidth(3, 200);
    reportSheet.getRange('D5').setValue('English').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(4, 30);
    reportSheet.getRange('E5').setValue('Training').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(5, 30);
    reportSheet.getRange('F5').setValue('Sales').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(6, 30);
    reportSheet.getRange('G5').setValue('CV').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(7, 150);
    reportSheet.getRange('H5').setValue('Upwork').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(8, 150);
    reportSheet.getRange('I5').setValue('Stack').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(9, 150);
    reportSheet.getRange('J5').setValue('Обучаемость\nСтрессоустойчивость\nРабота в команде\nРабота с клиентом\nНавыки самопрезентации\nГибкость мышления').setTextRotation(90).setBackground("#ffffff").setHorizontalAlignment("center").setVerticalAlignment("middle");
    reportSheet.setColumnWidth(10, 150);


    let column = 11;
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

    reportSheet.getRange(5,column).setValue('Plan').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(column, 150);
    column++;
    reportSheet.getRange(5,column).setValue('Fact').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(column, 150);
    column++;

    let columnForTable = column;
    let projects = getProjects(workloadSheet, null, true);

    // Write TOTAL in the next two columns
    reportSheet.getRange(5, column)
        .setValue('TOTAL plan')
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setTextRotation(90)
        .setBackground("#ffffff")
        .setFontSize(9);
    reportSheet.setColumnWidth(column, 40);

    // Leave a column for 'fact' data
    reportSheet.getRange(5, column + 1)
        .setValue('TOTAL fact')
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setTextRotation(90)
        .setBackground("#ffffff")
        .setFontSize(9);
    reportSheet.setColumnWidth(column + 1, 40);
    reportSheet.setColumnWidth(column + 2, 40);

    // Add a border to the right of the empty column
    reportSheet.getRange(5, column + 2, 120).setBorder(false, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);


    column = column+3;

    if(all) {


        for (let project of projects) {
            // Write the project name and PM initials in the next two columns
            reportSheet.getRange(5, column)
                .setValue(project.pmInitials + ' ' + project.projectName + ' plan')
                .setVerticalAlignment("middle")
                .setHorizontalAlignment("center")
                .setTextRotation(90)
                .setBackground("#cccccc")
                .setFontSize(9);
            reportSheet.setColumnWidth(column, 40);

            // Leave a column for 'fact' data
            reportSheet.getRange(5, column + 1)
                .setValue(project.pmInitials + ' ' + project.projectName + ' fact')
                .setVerticalAlignment("middle")
                .setHorizontalAlignment("center")
                .setTextRotation(90)
                .setBackground("#cccccc")
                .setFontSize(9);
            reportSheet.setColumnWidth(column + 1, 40);

            // Skip an empty column
            reportSheet.getRange(5, column + 2).setBackground("#ffffff");
            reportSheet.setColumnWidth(column + 2, 40);

            // Add a border to the right of the empty column
            reportSheet.getRange(5, column + 2, 120).setBorder(false, false, false, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

            // Increment the column counter to skip the 'fact' column
            column += 3;
        }
    }

    // Initialize the data rows
    let row = 7;
    for (let developer of developers) {
        if(!developer.name) continue;
        let developerName = developer.name.split("(")[0].trim(); // Remove everything after the "(" and trim spaces

        let allocationList = '';
        let developerAllocationData = allAllocationData[developerName];
        if (developerAllocationData && developerAllocationData.list) {
            allocationList = developerAllocationData.list;
        }

        //let competenceData = developerCompetenceData(developerName)
        let englishLevel = getCompetenceData(developerName)['Английский'];
        // Here you need to calculate trainingAndSales and allocation for each developer
        let trainingHours = developer.projects['Training'] || 0;
        let salesHours = developer.projects['SALES'] || 0;

        let stackData = getStackData(developerName);

        var developerCvData = getDeveloperCvList(developerName);
        var developerCvList = developerCvData.cvList;
        var developerUpworkData = getDeveloperUpworkData(developerName);
        var candidateData = getCandidateData(developerName);

        // Создание RichTextValueBuilder
        var combinedRichTextBuilder = SpreadsheetApp.newRichTextValue();
        var combinedText = '';

        // Добавление ссылки на папку CV
        if(developerCvData.folderId) {
            var cvFolderLink = 'https://drive.google.com/drive/folders/' + developerCvData.folderId;
            combinedText += 'CV folder\n\n';
        }

        // Добавление текста для каждого CV
        developerCvList.forEach(function(cv) {
            var date = Utilities.formatDate(new Date(cv.lastUpdate), 'GMT', 'dd/MM/yyyy');  // Преобразование даты в формат dd/MM/yyyy
            var linkText = cv.fileName + '\n';  // Текст ссылки
            var text = date + '\n\n';  // Дата обновления
            combinedText += linkText + text;
        });

        let cvText = combinedText;

        combinedRichTextBuilder.setText(combinedText);  // Обновляем текст в RichTextValueBuilder

        // Добавление ссылок для каждого CV
        var index = 0;
        if(developerCvData.folderId) {
            combinedRichTextBuilder.setLinkUrl(index, 'CV folder'.length, cvFolderLink);  // Ссылка на папку CV
            index += 'CV folder\n\n'.length;
        }

        developerCvList.forEach(function(cv) {
            var linkText = cv.fileName + '\n';  // Текст ссылки
            var text = Utilities.formatDate(new Date(cv.lastUpdate), 'GMT', 'dd/MM/yyyy') + '\n\n';  // Текст с датой
            var fullText = linkText + text;  // Полный текст
            combinedRichTextBuilder.setLinkUrl(index, index + linkText.length - 1, cv.link + '/edit');  // Ссылка на CV
            index += fullText.length;
        });

        // Построение итогового RichTextValue
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



        if (trainingHours >= 10) {
            // Выделить строку зеленым цветом
            reportSheet.getRange(row, 2, 1, 9).setBackground("#d9ead3"); // Смените число 11 на число столбцов в вашей строке
        }

        reportSheet.getRange(row, 2).setValue(developerName).setVerticalAlignment("middle").setWrap(true);
        reportSheet.getRange(row, 2).setNote(candidateComment);
        reportSheet.getRange(row, 3).setValue(developer.location).setVerticalAlignment("middle").setWrap(true);
        reportSheet.getRange(row, 4).setValue(englishLevel).setVerticalAlignment("middle");
        reportSheet.getRange(row, 5).setValue(trainingHours).setVerticalAlignment("middle");
        reportSheet.getRange(row, 6).setValue(salesHours).setVerticalAlignment("middle");

        // let profileLink = getCompetenceData(developerName)['личное дело'] ?? '';
        // if (profileLink) {
        //   reportSheet.getRange(row, 7).setFormula(`=HYPERLINK("${profileLink}", "Link")`).setVerticalAlignment("middle");
        // }

        if (cvDataRichText) {
            // Присваивание RichTextValue ячейке
            reportSheet.getRange(row, 7).setRichTextValue(cvDataRichText).setVerticalAlignment("middle").setHorizontalAlignment("left").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        }

        if (developerUpworkData.upworkLink) {
            // Присваивание обычного текста ячейке
            reportSheet.getRange(row, 8).setRichTextValue(upworkRichText).setVerticalAlignment("middle").setHorizontalAlignment("left").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
            reportSheet.getRange(row, 8).setNote(upworkText);
        }



        let competenceText = getCompetenceData(developerName)['Инструменты\nБиблиотеки\nСитстемы'] ?? '';
        reportSheet.getRange(row, 9).setValue('Main\n' + (getCompetenceData(developerName)['Основной стек'] ?? '') + '\n\nExtra\n' + (getCompetenceData(developerName)['Дополнительный стек'] ?? '')).setNote(competenceText).setVerticalAlignment("middle");
        reportSheet.getRange(row, 10).setValue(
            (getCompetenceData(developerName)['Обучаемость'] ?? '') + '  ' +
            (getCompetenceData(developerName)['Стрессоустойчивость'] ?? '') + '  ' +
            (getCompetenceData(developerName)['Работа в команде'] ?? '') + '  ' +
            (getCompetenceData(developerName)['Работа с клиентом (командой клиента)'] ?? '') + '  ' +
            (getCompetenceData(developerName)['Навыки самопрезентации'] ?? '') + '  ' +
            (getCompetenceData(developerName)['Гибкость мышления'] ?? '')
        ).setVerticalAlignment("middle").setHorizontalAlignment("center");

        let column = 11;
        n = 0;

        for (let stack of sortedStacks) {
            n++;
            let stackLevel = stackData[stack] || '';
            let cell = reportSheet.getRange(row, column);
            cell.setValue(stackLevel).setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");

            // Устанавливаем цвета для разных уровней
            if (stackLevel.startsWith('jun')) {
                cell.setBackground("#add8e6");  // Цвет для Junior (Светло-синий)
            } else if (stackLevel.startsWith('mid')) {
                cell.setBackground("#90ee90");  // Цвет для Middle (Светло-зелёный)
            } else if (stackLevel.startsWith('sr')) {
                cell.setBackground("#f4a460");  // Цвет для Senior (Светло-коричневый)
            }
            column++;
        }

        reportSheet.getRange(row, column).setValue(developer.projectHours).setVerticalAlignment("middle").setWrap(true);
        reportSheet.getRange(row, column+1).setValue(allocationList).setVerticalAlignment("middle").setWrap(true);

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

            if(all) {
                dataRow.push({
                    value: planHours,
                    verticalAlignment: "middle",
                    horizontalAlignment: "center",
                    background: "#cccccc",
                    fontSize: 8
                });

                dataRow.push({
                    value: factHours,
                    verticalAlignment: "middle",
                    horizontalAlignment: "center",
                    background: "#cccccc",
                    fontSize: 8
                });

                // Calculate the difference (plan - fact) and write in the next column
                let formula = `=IF(AND(ISBLANK(R${row}C${column}), ISBLANK(R${row}C${column+1})), "", R${row}C${column+1}-R${row}C${column})`;
                let difference = (factHours - planHours) || '';
                let color = difference < 0 ? "red" : "green";
                dataRow.push({
                    value: difference,
                    formula: formula,
                    verticalAlignment: "middle",
                    horizontalAlignment: "center",
                    background: "#ffffff",
                    fontSize: 8,
                    fontColor: color
                });

                // Write all the data to the row at once
                let range = reportSheet.getRange(row, column, 1, 3);
                range.setValues([dataRow.map(cell => cell.value)]);
                range.setBackgrounds([dataRow.map(cell => cell.background)]);
                range.setFontColors([dataRow.map(cell => cell.fontColor)]);
                range.setFontSizes([dataRow.map(cell => cell.fontSize)]);
                range.setVerticalAlignments([dataRow.map(cell => cell.verticalAlignment)]);
                range.setHorizontalAlignments([dataRow.map(cell => cell.horizontalAlignment)]);

                // Skip an empty column
                reportSheet.getRange(5, column + 2).setBackground("#ffffff");
                reportSheet.setColumnWidth(column + 2, 35);

                // Increment the column counter to skip the 'fact' column
                column += 3;
            }

            planHours = Math.round(planHours * 100) / 100;
            factHours = Math.round(factHours * 100) / 100;

            planHoursTotal += planHours;
            factHoursTotal += factHours;

        }

        let diffHoursTotal = factHoursTotal-planHoursTotal;

        let diffFontColor = "green"
        if(diffHoursTotal < 0) diffFontColor = "red";

        reportSheet.getRange(row, columnPlanHours).setValue(planHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center");
        reportSheet.getRange(row, columnFactHours).setValue(factHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center");
        reportSheet.getRange(row, columnDiffHours).setValue(diffHoursTotal).setVerticalAlignment("middle").setHorizontalAlignment("center").setFontColor(diffFontColor);

        row++;
        reportSheet.setRowHeight(row, 150);
    }

    // Set the border
    reportSheet.getRange(5, 2, row-5, columnForTable-2).setBorder(true, true, true, true, true, true);

    insertSumFormulas(all,isLastWeek);
    let lastColumn = reportSheet.getLastColumn();

    // определите номер строки, куда нужно вставить итоговые значения (после последней строки с данными)
    let totalRow = reportSheet.getLastRow() + 1;

    reportSheet.getRange(4, 10).setValue('доступные ресурсы:');

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

    // Apply the first formula from startColumn to endColumn
    for(let i = 11; i <= endColumn; i++) {
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

    // создаем новый фильтр
    var range = reportSheet.getRange(6, 2, sheet.getLastRow() - 5, lastColumn-1);
    range.createFilter();

    // Add date and time of data gathering
    const currentTime = new Date().toLocaleString("en-GB", {timeZone: "Asia/Tbilisi"});
    reportSheet.getRange("B4").setValue(`Generated at ${currentTime} (Tbilisi, Georgia Timezone)`);

    // Определите диапазон строк, которому вы хотите задать новую высоту
    var startRow = 6;
    var numRows = reportSheet.getLastRow() - startRow + 1;

    // Установите высоту всех строк в этом диапазоне
    reportSheet.setRowHeights(startRow, numRows, 50);

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

function getDeveloperStackData(developerName) {
    let data = getDeveloperCompetenceData(developerName);

    let mainStack = data['Основной стек'] ? data['Основной стек'].split('\n') : [];
    let extraStack = data['Дополнительный стек'] ? data['Дополнительный стек'].split('\n') : [];
    let stackData = {};

    mainStack.concat(extraStack).forEach(stackLine => {
        let lastSpaceIndex = stackLine.lastIndexOf(' ');
        let stack = stackLine.substring(0, lastSpaceIndex).trim();
        let level = stackLine.substring(lastSpaceIndex + 1);

        if (level) {
            level = level.toLowerCase().trim().replace('middle', 'mid').replace('junior', 'jun').replace('senior', 'sr').replace('?', '').replace('nonchecked', '');
        }
        if (level != 'nonchecked') stackData[stack] = level || '';
    });

    return stackData;
}

function getDeveloperStackDataFromSheet(developerNameRussian) {
    let namesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Developers english vs russian names');
    let namesData = namesSheet.getRange(2, 1, namesSheet.getLastRow(), 2).getValues();

    let nameMatch = namesData.find(row => row[1] === developerNameRussian);
    if (!nameMatch) {
        Logger.log(`Разработчик с именем ${developerNameRussian} не найден в листе 'Developers english vs russian names'`);
        return {};
    }

    let developerName = nameMatch[0]; // Получить имя на английском

    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeveloperStackData');
    let data = sheet.getRange(2, 1, sheet.getLastRow(), 6).getValues();
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

    return stackData;
}

function getAllDevelopersStackDataFromSheet() {
    let namesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Developers english vs russian names');
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


function getAllocationData(developers, isLastWeek = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let scrumSheetName = isLastWeek ? 'Scrum files for last week' : 'Scrum files for current week';
    const scrumSheet = ss.getSheetByName(scrumSheetName);

    let allocationData = {};
    let scrumData;

    try {
        const range = scrumSheet.getRange('A3:E' + scrumSheet.getLastRow());
        scrumData = range.getValues();
    } catch (error) {
        SpreadsheetApp.getUi().alert("Error retrieving data from the scrum sheet: " + error);
        return;
    }

    scrumData.forEach(row => {
        if (row[2] == "HR") row[3] = "HR";
        if (row[2] == "PRESALE") row[3] = "SALES";
        if (row[2] == "Administrative") row[3] = "Administrative";
        if (row[2] == "Testing") row[3] = "Testing";
        if (row[2] == "DevOps") row[3] = "DevOps";
        const [developerShort, date, type, project, hours] = row;
        let roundedHours = Math.round(hours * 100) / 100;
        let developerFull = developers.find(developer => developer.name.startsWith(developerShort));
        if (developerFull && developerFull.name) {
            developerFull.name = developerFull.name.split("(")[0].trim();
            if (!allocationData[developerFull.name]) {
                allocationData[developerFull.name] = {projects: {}, list: ''};
            }
            if (!allocationData[developerFull.name].projects[project]) {
                allocationData[developerFull.name].projects[project] = 0;
            }
            allocationData[developerFull.name].projects[project] += roundedHours;
            Logger.log(developerFull.name + ' ' + project + ' ' + roundedHours);
        }
    });

    for (let developer in allocationData) {
        let allocationList = [];
        let totalDeveloperHours = 0;
        for (let project in allocationData[developer].projects) {
            let roundedHours = Math.round(allocationData[developer].projects[project] * 100) / 100;
            totalDeveloperHours += roundedHours;
            allocationList.push(project + ' (' + roundedHours + ')');
            Logger.log(project + ' (' + roundedHours + ')');
        }
        allocationList.unshift(Math.round(totalDeveloperHours * 100) / 100 + '');
        allocationData[developer].list = allocationList.join(' | ');
    }

    return allocationData;
}



function getDevelopers(workloadSheet, all) {
    if (!workloadSheet) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const workloadSheetId = "1N65NUtqBA855C6K8swmeFQ9HbvIZU4fq4EnhYzvNV7Q";
        const workloadSpreadsheet = SpreadsheetApp.openById(workloadSheetId);

        let weekMondayDate = getCurrentMonday();
        let weekSundayDate = getCurrentSunday();

        const weekMondayString = Utilities.formatDate(weekMondayDate, ss.getSpreadsheetTimeZone(), 'd MMM').toLowerCase();
        const weekSundayString = Utilities.formatDate(weekSundayDate, ss.getSpreadsheetTimeZone(), 'd MMM').toLowerCase();

        let workloadSheetName = weekMondayDate.getMonth() === weekSundayDate.getMonth() ?
            `${weekMondayString.split(" ")[0]}-${weekSundayString.split(" ")[0]} ${weekSundayString.split(" ")[1]}` :
            `${weekMondayString}-${weekSundayString.split(" ")[0]}`;

        workloadSheet = workloadSpreadsheet.getSheetByName(workloadSheetName);
        if (!workloadSheet) {
            SpreadsheetApp.getUi().alert(`Cannot find sheet "${workloadSheetName}" in the workload spreadsheet.`);
            return;
        }
    }

    let developers = [];
    let projects = [];

    let workloadData = workloadSheet.getDataRange().getValues();

    // Retrieve projects from the 5th row
    projects = workloadData[4].slice(1);

    // Iterate through the rows of the workloadData
    for (let i = 5; i < workloadData.length; i++) {
        // Get the developer's name, which is assumed to be in the 4th column
        let developerName = workloadData[i][3];
        let developerLocation = workloadData[i][0];
        developerName = developerName.split("(")[0].trim();

        Logger.log(developerName);

        var projectHours = getHoursByNameAndProject(workloadData, developerName);

        // If the developer name is "total", stop the loop
        if (developerName === 'total') {
            break;
        }

        if (!developerName) {
            continue;
        }

        // Create a new developer object
        let developer = {
            name: developerName,
            location: developerLocation,
            projectHours,
            projects: {},
            vacationHours: workloadData[i][projects.indexOf('vacation') + 1] || 0,  // Add vacation hours
        };

        let workedOnTraining = false;
        let workedOnSales = false;

        for (let j = 5; j < workloadData[i].length; j++) {
            hours = workloadData[i][j] || 0;
            let projectName = projects[j - 1] || "(noname)";

            if (hours>0) {
                developer.projects[projectName] = hours;
                if (projectName == "Training") {
                    workedOnTraining = true;
                } else if (projectName == "SALES") {
                    workedOnSales = true;
                }
            }
        }

        if (all || (workedOnTraining || workedOnSales)) {
            developers.push(developer);
        }
    }

    return developers;
}


function getProjects(workloadSheet, projectNameFilter, isLastWeek = false) {
    if(!workloadSheet) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const workloadSheetId = "1N65NUtqBA855C6K8swmeFQ9HbvIZU4fq4EnhYzvNV7Q";
        const workloadSpreadsheet = SpreadsheetApp.openById(workloadSheetId);

        let weekMondayDate;
        let weekSundayDate;

        if (isLastWeek) {
            weekMondayDate = getLastMonday();
            weekSundayDate = getLastSunday();
        } else {
            weekMondayDate = getCurrentMonday();
            weekSundayDate = getCurrentSunday();
        }

        const weekMondayString = Utilities.formatDate(weekMondayDate, ss.getSpreadsheetTimeZone(), 'd MMM').toLowerCase();
        const weekSundayString = Utilities.formatDate(weekSundayDate, ss.getSpreadsheetTimeZone(), 'd MMM').toLowerCase();

        let workloadSheetName = weekMondayDate.getMonth() === weekSundayDate.getMonth() ?
            `${weekMondayString.split(" ")[0]}-${weekSundayString.split(" ")[0]} ${weekSundayString.split(" ")[1]}` :
            `${weekMondayString}-${weekSundayString.split(" ")[0]}`;

        workloadSheet = workloadSpreadsheet.getSheetByName(workloadSheetName);
        if (!workloadSheet) {
            SpreadsheetApp.getUi().alert(`Cannot find sheet "${workloadSheetName}" in the workload spreadsheet.`);
            return;
        }
    }

    let projects = [];
    let developers = [];

    let workloadData = workloadSheet.getDataRange().getValues();

    // Retrieve developers from the 4th column
    developers = workloadData.map(row => row[3]);

    // Iterate through the columns of the workloadData
    outer: for (let j = 1; j < workloadData[0].length; j++) {
        // Get the project's name, which is assumed to be in the 5th row
        let projectName = workloadData[4][j].trim() || "(noname)";
        console.log('Project Name:', projectName);

        // Get the PM's initials, which are assumed to be in the 1st row
        let pmInitials = workloadData[0][j].trim();
        console.log('PM Initials:', pmInitials);

        // If PM initials are not exactly 2 characters long or not Russian letters, skip this iteration
        let regex = /[а-яё]{2}/i;
        if (pmInitials === undefined || pmInitials === null || pmInitials.length !== 2 || !regex.test(pmInitials)) {
            continue outer;
        }

        // If a projectNameFilter is provided and doesn't match the current project, skip this iteration
        if (projectNameFilter && projectName !== projectNameFilter) {
            continue outer;
        }

        // Create a new project object
        let project = {pmInitials, projectName, projectHours: 0, developers: {}};

        for (let i = 5; i < workloadData.length; i++) {
            let hours = workloadData[i][j] || 0;
            let developerName = developers[i];

            if (developerName === 'total') {
                break;
            }

            developerName = developerName.split("(")[0].trim();

            if (hours > 0) {
                project.developers[developerName] = hours;
                project.projectHours += hours;
            }
        }

        projects.push(project);
    }

    Logger.log(projects);
    return projects;
}

function getScrumFilesData(fromDate, toDate) {
    const spreadsheetId = "133dteyNbEFrZgxxIDnI3CytGRGyk_t6U3uUpgNAN0nc";
    const sheetName = "Scrum files";
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const scrumFilesSheet = ss.getSheetByName(sheetName);

    // Retrieve list of scrum files urls starting from the 3rd row of the 'B' column
    const scrumFilesUrls = scrumFilesSheet.getRange('B3:B' + scrumFilesSheet.getLastRow()).getValues().flat();

    if(scrumFilesUrls.length == 0) {
        Logger.log('No matching files found.');
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

    Logger.log(JSON.stringify(data)); // Log the entire data object
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


function getLastMonday() {
    const currentMonday = getCurrentMonday();
    const lastMonday = new Date(currentMonday.getFullYear(), currentMonday.getMonth(), currentMonday.getDate() - 7);
    return lastMonday;
}


function getLastSunday() {
    const currentSunday = getCurrentSunday();
    const lastSunday = new Date(currentSunday.getFullYear(), currentSunday.getMonth(), currentSunday.getDate() - 7);
    return lastSunday;
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

function gatherDataInSheet(isLastWeek) {
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

function copyDataToCompetencesSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let competencesSheet = ss.getSheetByName('Competences');

    // Если лист Competences не существует, создайте его
    if (!competencesSheet) {
        competencesSheet = ss.insertSheet('Competences');
    }

    // Откройте другой документ и получите данные
    const sourceSpreadsheetId = '1T25tKDuj3DqAJcX1PE8lIVhr62kN0IaVZWfzAYpcJUc';
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    const dataSheet = sourceSpreadsheet.getSheetByName('Data');
    const sourceData = dataSheet.getRange('A4:N' + dataSheet.getLastRow()).getValues();

    // Очистите лист Competences и вставьте новые данные
    competencesSheet.clear();

    // Add date and time of data gathering
    const currentTime = new Date().toLocaleString("en-GB", {timeZone: "Asia/Tbilisi"});
    competencesSheet.getRange("A1").setValue(`Data gathered at ${currentTime} (Tbilisi, Georgia Timezone)`);

    competencesSheet.getRange(2, 1, sourceData.length, sourceData[0].length).setValues(sourceData);

    // Установите высоту всех строк в 150
    competencesSheet.setRowHeights(1, sourceData.length, 150);
}

function collectCandidatesData() {
    var sourceDocId = '189YZ_AKtBhVBADGksYIjKQCg8h_ky6Bh5tjEzxUWeXY';
    var sourceSheetName = 'Candidates';
    var destSheetName = 'CandidatesData';

    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var developersSheet = activeSpreadsheet.getSheetByName('Developers english vs russian names');
    var developersRange = developersSheet.getDataRange();
    var developersValues = developersRange.getValues().map(row => row[1]).slice(1);

    var developersNames = developersValues.flat();  // Получаем список разработчиков
    Logger.log("Developers Names: " + developersNames);

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

    Logger.log("Headers written to destination sheet");

    // Записываем данные кандидатов
    for (var sourceRowIndex = 1; sourceRowIndex < sourceValues.length; sourceRowIndex++) {
        var row = sourceValues[sourceRowIndex];
        var name = row[0];
        var status = row[2];  // Исправлен индекс на 2, так как в колонке C статус

        if (developersNames.includes(name) && status === 'Принят') {
            Logger.log("Found matching candidate: " + name);

            // Добавляем ID в конец строки
            row.push(sourceRowIndex + 1);

            // Получаем ссылки из колонок L и AJ
            let linkRichTextL = sourceRichTextValues[sourceRowIndex][headers.indexOf('L')];  // Get the RichTextValue from column L
            if (linkRichTextL) {
                let linkUrlL = linkRichTextL.getLinkUrl();
                if (linkUrlL) {
                    row[headers.indexOf('L')] = linkUrlL;
                } else {
                    Logger.log("Failed to get link from column L");
                }
            }

            let linkRichTextAJ = sourceRichTextValues[sourceRowIndex][headers.indexOf('AJ')];  // Get the RichTextValue from column AJ
            if (linkRichTextAJ) {
                let linkUrlAJ = linkRichTextAJ.getLinkUrl();
                if (linkUrlAJ) {
                    row[headers.indexOf('AJ')] = linkUrlAJ;
                } else {
                    Logger.log("Failed to get link from column AJ");
                }
            }

            destSheet.appendRow(row);
        }
    }



    // Записываем время генерации данных
    var now = Utilities.formatDate(new Date(), 'Asia/Tbilisi', 'dd/MM/yyyy, HH:mm:ss');
    destSheet.getRange('A1').setValue('Generated at ' + now + ' (Tbilisi, Georgia Timezone)');

    Logger.log('Process completed');
}

function collectDeveloperCvData() {
    var folderId = '0B5SXKqmca-G9azQtSHN6YTlKOUU';
    var destSheetName = 'DeveloperCvData';

    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var destSheet = activeSpreadsheet.getSheetByName(destSheetName);
    if (!destSheet) {
        // Если лист не существует, создаем новый
        destSheet = activeSpreadsheet.insertSheet(destSheetName);
    }

    // Получаем имена разработчиков с листа "Developers english vs russian names"
    var developersSheet = activeSpreadsheet.getSheetByName('Developers english vs russian names');
    var developersValues = developersSheet.getDataRange().getValues().slice(1);  // Exclude headers
    var englishNames = developersValues.map(row => row[0]);
    var russianNames = developersValues.map(row => row[1]);

    // Очищаем лист
    destSheet.clear();

    // Записываем заголовки
    var headers = ['Имя папки', 'ID папки', 'Имя файла', 'ID файла', 'Дата последнего изменения файла', 'CV link'];
    destSheet.getRange('A2:' + columnToLetter(headers.length) + '2').setValues([headers]);

    Logger.log("Headers written to destination sheet");

    var folder = DriveApp.getFolderById(folderId);
    var subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
        var subFolder = subFolders.next();
        var folderName = subFolder.getName();
        Logger.log('Processing folder: ' + folderName);

        // Проверяем, что имя папки соответствует русскому или английскому имени разработчика
        if (englishNames.includes(folderName) || russianNames.includes(folderName)) {
            var files = subFolder.getFiles();
            while (files.hasNext()) {
                var file = files.next();
                Logger.log('Processing file: ' + file.getName());

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
        } else {
            Logger.log('Folder name does not match any developer names: ' + folderName);
        }
    }

    // Записываем время генерации данных
    var now = Utilities.formatDate(new Date(), 'Asia/Tbilisi', 'dd/MM/yyyy, HH:mm:ss');
    destSheet.getRange('A1').setValue('Generated at ' + now + ' (Tbilisi, Georgia Timezone)');

    Logger.log('Process completed');
}


function collectDeveloperUpworkData() {
    var sourceDocId = '1arJRaIn_0B-0ds32JkccZrUxbZ_GlSbT4imDoZU8iKo';
    var sourceSheetName = 'All';
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
    headers.push('Имя агентства', 'upwork-account');
    destSheet.getRange('A2:' + columnToLetter(headers.length) + '2').setValues([headers]);

    Logger.log("Headers written to destination sheet");

    var agencyName = '';
    for(var i = 1; i < sourceData.length; i++) {
        var row = sourceData[i];

        // Проверка, является ли строка строкой агентства
        if(row.slice(1).every(cell => !cell)) {
            agencyName = row[0]; // Обновляем имя агентства
        } else {
            var upworkAccountMatch = row[0].match(/\(([^)]+)\)/); // Ищем upwork-account в скобках
            var upworkAccount = upworkAccountMatch ? upworkAccountMatch[1] : '';
            row.push(agencyName, upworkAccount); // Добавляем имя агентства и upwork-account в строку
            destSheet.appendRow(row);
        }
    }

    // Записываем время генерации данных
    var now = Utilities.formatDate(new Date(), 'Asia/Tbilisi', 'dd/MM/yyyy, HH:mm:ss');
    destSheet.getRange('A1').setValue('Generated at ' + now + ' (Tbilisi, Georgia Timezone)');

    Logger.log('Process completed');
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nameTranslationsSheet = ss.getSheetByName('Developers english vs russian names');

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


function getDeveloperCompetenceData(developerName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const competencesSheet = ss.getSheetByName('Competences');

    if (!competencesSheet) {
        Logger.log(`Sheet "Competences" not found.`);
        return;
    }

    const competencesData = competencesSheet.getDataRange().getValues();
    const headers = competencesData[1]; // First row contains the headers
    let developerCompetenceData = {};

    // Loop through the rest of the rows to find the matching developer
    for (let i = 2; i < competencesData.length; i++) {
        if (nameTranslations[competencesData[i][0].trim()] === developerName.trim()) {  // If the developer name in the first column matches
            // Loop through the rest of the columns for this row
            for (let j = 0; j < headers.length; j++) {
                let header = headers[j];
                let value = competencesData[i][j];
                developerCompetenceData[header] = value;
            }
            break;
        }
    }
    return developerCompetenceData;
}



function getAllDevelopersCompetenceData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const competencesSheet = ss.getSheetByName('Competences');

    if (!competencesSheet) {
        Logger.log(`Sheet "Competences" not found.`);
        return;
    }

    const namesSheet = ss.getSheetByName('Developers english vs russian names');
    const namesData = namesSheet.getRange(2, 1, namesSheet.getLastRow(), 2).getValues();
    const nameTranslations = Object.fromEntries(namesData.map(row => [row[0], row[1]]));

    const competencesData = competencesSheet.getDataRange().getValues();
    const headers = competencesData[1]; // First row contains the headers

    let allDevelopersCompetenceData = {};

    // Loop through the rest of the rows to find the matching developer
    for (let i = 2; i < competencesData.length; i++) {
        let developerData = {};
        let developerNameEnglish = competencesData[i][0].trim();
        let developerNameRussian = nameTranslations[developerNameEnglish];

        // Loop through the rest of the columns for this row
        for (let j = 1; j < headers.length; j++) {
            let header = headers[j];
            let value = competencesData[i][j];
            developerData[header] = value;
        }

        allDevelopersCompetenceData[developerNameRussian] = developerData;
    }

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

        var developerName = row[0]; // Имя разработчика
        var folderId = row[1]; // ID папки
        var cvFileName = row[2]; // Имя файла CV
        var cvFileId = row[3]; // ID файла
        var lastUpdate = row[4]; // Дата обновления файла
        var cvLink = row[5]; // Ссылка на резюме

        if (!developersCvData[developerName]) {
            developersCvData[developerName] = { folderId: folderId, cvList: [] };
        }

        developersCvData[developerName].cvList.push({ fileName: cvFileName, fileId: cvFileId, link: cvLink, lastUpdate: lastUpdate });
    });

    return developersCvData;
}

function getAllDevelopersUpworkDataFromSheet() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeveloperUpworkData');
    var data = sheet.getDataRange().getValues();
    var developersUpworkData = {};

    // Получаем данные о соответствии имен с листа "Developers english vs russian names"
    var namesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Developers english vs russian names');
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
            var upworkLink = row[6];
            var upworkAccount = row[0];
            var agencyName = row[20];
            var busyHours = row[1];
            var busyProjects = row[2];
            var RP = row[3];
            var classifier = row[4];
            var rate2021 = row[5];

            developersUpworkData[russianName] = {
                upworkLink: upworkLink,
                upworkAccount: upworkAccount,
                agencyName: agencyName,
                busyHours: busyHours,
                busyProjects: busyProjects,
                RP: RP,
                classifier: classifier,
                rate2021: rate2021,
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


function getHoursByNameAndProject(data, name) {
    var hoursAndProjects = [];
    for (var i = 0; i < data.length; i++) {
        var rowName = data[i][3].toString();
        //Logger.log(data[i][3]);
        if (rowName.startsWith(name)) {
            for (var j = 5; j < data[0].length; j++) {
                var cellValue = data[i][j];
                var hours = parseFloat(cellValue);
                if (hours > 0) {
                    var pm = data[0][j];
                    var project = data[4][j];
                    project = project.trim();
                    if (pm == '') break;
                    hoursAndProjects.push(pm + " " + project + " (" + hours.toFixed(2) + ")");
                }
            }
            break;
        }
    }
    return hoursAndProjects.join(' | ');
}

function updateDeveloperStackData() {
    Logger.log('Начало обработки файлов');

    var folder = DriveApp.getFolderById('15cKH1ynPdkLLv-UXLToNUCy8QrfOCTXm');
    var files = folder.getFiles();
    var output = [];
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('DeveloperStackData');

    if (sheet == null) {
        sheet = spreadsheet.insertSheet('DeveloperStackData');
        Logger.log('Лист "DeveloperStackData" создан');
    } else {
        sheet.clear();
        Logger.log('Лист "DeveloperStackData" очищен');
    }

    while (files.hasNext()) {
        var file = files.next();
        var fileName = file.getName();

        if (fileName.toLowerCase().includes('personal')) {
            Logger.log('Пропущен файл: ' + fileName);
            continue;
        }

        Logger.log('Обработка файла: ' + fileName);
        var fileSpreadsheet = SpreadsheetApp.openById(file.getId());
        var fileSheet = fileSpreadsheet.getSheets()[0];

        output.push(...getDataFromRange(fileSheet, 'A:D', 'Основной', fileName));
        output.push(...getDataFromRange(fileSheet, 'E:H', 'Дополнительный', fileName));
    }

    // Add date and time of data gathering
    const currentTime = new Date().toLocaleString("en-GB", {timeZone: "Asia/Tbilisi"});
    sheet.appendRow([`Data gathered at ${currentTime} (Tbilisi, Georgia Timezone)`]);

    sheet.appendRow(['Имя разработчика', 'Тип', 'Технология', 'Уровень', 'Желание/Нежелание', 'Стек']);

    sheet.getRange(3, 1, output.length, output[0].length).setValues(output);
    correctDataInDeveloperStackData();
    Logger.log('Данные записаны в лист');
}

function getDataFromRange(sheet, range, stackType, developerName) {
    Logger.log('Сбор данных из диапазона: ' + range);

    var data = sheet.getRange(range).getValues();
    var output = [];

    for (var i = 4; i < data.length; i++) {
        var row = data[i];
        if (row[0] === '') {
            break;
        }
        output.push([developerName, ...row, stackType]);
    }

    Logger.log('Данные из диапазона ' + range + ' собраны');
    return output;
}

function correctDataInDeveloperStackData() {
    Logger.log('Начало корректировки данных');

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeveloperStackData');
    var data = sheet.getRange(3, 1, sheet.getLastRow(), 6).getValues();
    var correctedData = [];

    var types = [...new Set(data.map(row => row[1].toLowerCase()))];
    var technologies = [...new Set(data.map(row => row[2].toLowerCase()))];
    //var levels = [...new Set(data.map(row => row[3]))]; // Здесь мы не приводим к нижнему регистру

    Logger.log('Получение уникальных значений для типов, технологий и уровней');

    var typeCorrections = getCorrections(types, data, 1);
    var technologyCorrections = getCorrections(technologies, data, 2, true, 1); // Установить порог схожести в 1 для технологий
    //var levelCorrections = getCorrections(levels, data, 3, false); // Не применяем toLowerCase для уровней

    Logger.log('Корректировка данных');

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

    Logger.log('Завершено: данные скорректированы и обновлены на листе');
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Получаем оба листа
    var sheet1 = ss.getSheetByName('ALL report');
    var sheet2 = ss.getSheetByName('ALL report last week');

    // Массив для всех слов
    var allWords = [];

    // Функция для сбора данных с определенного листа
    function gatherDataFromSheet(sheet) {
        var dataRange = sheet.getDataRange();
        var values = dataRange.getValues();

        // Получаем все комментарии из листа
        var notes = dataRange.getNotes();

        values.forEach(function(row) {
            row.forEach(function(cell) {
                var words = cell.toString().replace(/\n/g, " ").toLowerCase().replace(/[0-9]+|[^\wа-яА-Я_+.-]+/g, " ").split(" ");
                allWords = allWords.concat(words);
            });
        });

        notes.forEach(function(noteRow) {
            noteRow.forEach(function(note) {
                var noteWords = note.toString().replace(/\n/g, " ").toLowerCase().replace(/[0-9]+|[^\wа-яА-Я_+.-]+/g, " ").split(" ");
                allWords = allWords.concat(noteWords);
            });
        });
    }

    // Собираем данные с каждого листа
    gatherDataFromSheet(sheet1);
    gatherDataFromSheet(sheet2);

    var uniqueWords = Array.from(new Set(allWords));
    uniqueWords = uniqueWords.filter(word => word.length > 1 && word[0] !== '+' && word[0] !== '-');
    uniqueWords.sort();

    // Создаем лист "Keywords", если его еще нет
    var keywordSheet = ss.getSheetByName('Keywords');
    if (!keywordSheet) {
        keywordSheet = ss.insertSheet('Keywords');
    }

    keywordSheet.clear();

    for (var i = 0; i < uniqueWords.length; i++) {
        keywordSheet.getRange(i + 1, 1).setValue(uniqueWords[i]);
    }
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

function insertSumFormulas(all, isLastWeek = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheetName = "SALES report"
    if(all) {
        sheetName = 'ALL report';
    }

    if(isLastWeek) {
        sheetName += ' last week';
    }

    sheet = ss.getSheetByName(sheetName);

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

function testGetAllDevelopersUpworkDataFromSheet() {
    var developersUpworkData = getAllDevelopersUpworkDataFromSheet();
    Logger.log(JSON.stringify(developersUpworkData, null, 2));
}
