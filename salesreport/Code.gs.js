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
        // .addItem('Выбор проектов', 'showProjectsDialog')
        .addItem('Сортировать по A-Z', 'sortDataAscending')
        .addItem('Сортировать по Z-A', 'sortDataDescending')
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

    var dataRange = sheet.getRange(6, 1, sheet.getLastRow() - 5, sheet.getLastColumn());

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

    var dataRange = sheet.getRange(6, 1, sheet.getLastRow() - 5, sheet.getLastColumn());

    dataRange.sort([{column: sortColumn, ascending: false}]);
}

function getActiveStacks() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var headerRow = sheet.getRange("5:5");
    var cells = headerRow.getValues()[0];  // Изменено здесь
    var activeStacks = [];
    for (var i = 10; i < cells.length; i++) {
        var cell = sheet.getRange(5, i+1);   // Получаем ячейку для проверки ее свойств
        if (cell.getBackground() === '#000000') {
            activeStacks.push(cells[i]);
        }
        if (cell.getValue() === 'Plan') {
            break;
        }
    }
    return activeStacks;
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

function toggleStack(stackName) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var headerRow = sheet.getRange("5:5");
    var values = headerRow.getValues();
    var stackColumn = values[0].indexOf(stackName)+1;
    var totalPlanColumn = values[0].indexOf('Plan');

    if (stackColumn >= 10 && stackColumn < totalPlanColumn) {
        var cell = headerRow.getCell(1, stackColumn);
        var fontColor = cell.getFontColorObject().asRgbColor().asHexString();

        if (fontColor == '#ffffff') { // Если ячейка уже активна
            cell.setBackground('#cccccc').setFontColor('black');

            // Получить все активные стеки
            let activeStacks = getActiveStacks();

            // Вызовите sortData с активными стеками
            sortData(activeStacks);
        } else { // Если ячейка не активна
            cell.setBackground('#000000').setFontColor('white');
            hideEmptyRows(sheet, stackColumn); // Скрыть строки без значений для этого стека
        }
    } else {
        Logger.log("Не удалось найти стек " + stackName);
    }

    return true;
}

function sortData(activeStacks) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return false;
    }

    // Показываем все строки
    showAllRows();

    // Для каждого активного стека вызываем функцию toggleStack
    activeStacks.forEach(stackName => {
        toggleStack(stackName);
    });

    // Сбрасываем поиск
    sheet.getRange("H3").clearContent();
    sheet.getRange("I3").clearContent();

    return true;
}


function hideRowsForActiveStacks() {
    var activeStacks = getActiveStacks(); // Вставьте сюда код для получения активных стеков
    var sheet = SpreadsheetApp.getActiveSheet();
    var lastColumn = sheet.getLastColumn();
    var headerRow = sheet.getRange("5:5");
    var values = headerRow.getValues();

    activeStacks.forEach(function(stackName) {
        var stackColumn = values[0].indexOf(stackName)+1;
        if (stackColumn >= 10 && stackColumn <= lastColumn) {
            hideEmptyRows(sheet, stackColumn);
        }
    });
}

function filterByProject(projectName) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var lastRow = sheet.getLastRow();
    var dataRange = sheet.getRange(6, 6, lastRow - 5, 2); // Колонки F и G, начиная с 6-й строки и до конца
    var values = dataRange.getValues();

    // Показать все строки перед применением нового фильтра
    showAllRows();

    // Пройти через все строки и скрыть те, которые не содержат projectName
    var rowsToHide = [];
    for (var i = 0; i < values.length; i++) {
        if (!values[i][0].includes(projectName) && !values[i][1].includes(projectName)) {
            rowsToHide.push(i + 6); // Собираем номера строк для последующего скрытия
        }
    }
    // Скрываем все строки сразу
    for (var i = 0; i < rowsToHide.length; i++) {
        sheet.hideRows(rowsToHide[i]);
    }
    // Сбрасываем поиск
    sheet.getRange("H3").clearContent();
    sheet.getRange("I3").clearContent();
    sheet.getRange("H3").setValue("").setBackground('white').setFontColor('black');
    sheet.getRange("I3").setValue("").setBackground('white').setFontColor('black');
    sheet.getRange("J3").setValue("").setBackground('white').setFontColor('black');
}

function hideEmptyRows(sheet, sortColumn) {
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var dataRange = sheet.getRange(6, sortColumn, sheet.getLastRow() - 5);
    var data = dataRange.getValues();

    for (var i = 0; i < data.length; i++) {
        if (data[i][0] === "" || data[i][0] === null) {
            sheet.hideRows(i + 6); // Скрыть пустые строки
        }
    }
}

function searchData(query) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    showAllRows();

    var dataRange = sheet.getRange(6, 1, lastRow - 5, lastColumn); // Мы начинаем с 6-й строки и идем до последней строки
    var values = dataRange.getValues();
    sheet.showRows(1, values.length);

    var notesColumnK = sheet.getRange(6, 11, lastRow - 5).getNotes(); // Получаем заметки из колонки K (11)

    if(query) {
        sheet.getRange("H3").setValue("Поиск по запросу:").setBackground('black').setFontColor('white');
        sheet.getRange("I3").setValue(query).setBackground('black').setFontColor('white');
        sheet.getRange("J3").setBackground('black').setFontColor('white');
    } else {
        sheet.getRange("H3").setValue("").setBackground('white').setFontColor('black');
        sheet.getRange("I3").setValue("").setBackground('white').setFontColor('black');
        sheet.getRange("J3").setValue("").setBackground('white').setFontColor('black');
    }


    // Скрываем все строки в диапазоне данных
    sheet.hideRows(6, values.length);

    // Преобразуем запрос в массив слов
    var queryWords = query.toLowerCase().split(" ");

    for(var i = 0; i < values.length; i++) {
        // Если строка содержит все слова из запроса, показываем ее
        var rowContent = values[i].join(" ") + " " + notesColumnK[i][0]; // Добавляем содержимое заметки к строке
        var rowContentLower = rowContent.toLowerCase();

        var containsAllWords = queryWords.every(function(word) {
            return rowContentLower.includes(word);
        });

        if(containsAllWords) {
            sheet.showRows(i + 6); // Показываем строки, начиная с 6-й строки
        }
    }
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

    for (var i = 5; i < data.length; i++) {
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
    sheet.getRange("J3").setValue("").setBackground('white').setFontColor('black');
}

function showAllRows() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report' && sheetName !== 'ALL report last week') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report", "SALES report" или "ALL report last week".');
        return;
    }

    var data = sheet.getDataRange().getValues();
    sheet.showRows(1, data.length);

    // Reset the background color and text color of stack headers
    var headerRow = sheet.getRange("5:5");
    var headerValues = headerRow.getValues()[0];
    var lastColumn = sheet.getLastColumn();
    var backgrounds = headerRow.getBackgrounds()[0];
    var fontColors = headerValues.map((_, index) => {
        return headerRow.getCell(1, index + 1).getFontColorObject().asRgbColor().asHexString();
    });

    var totalPlanColumn = headerValues.indexOf("Plan");
    if (totalPlanColumn == -1) {
        totalPlanColumn = lastColumn;  // Если "Plan" не найден, используем последнюю колонку
    }

    for (var i = 10; i < totalPlanColumn; i++) {
        backgrounds[i] = '#cccccc';
        fontColors[i] = 'black';
    }

    // Заполняем все ячейки значениями по умолчанию
    headerRow.setBackgrounds([backgrounds]);
    headerRow.setFontColors([fontColors]);

    // Сбрасываем поиск
    sheet.getRange("H3:J3").clearContent().setBackground('white').setFontColor('black');

    return true;
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

    showAllRows();

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

    Logger.log(developers.length);

    showAllRows();

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
    reportSheet.getRange('G5').setValue('Profile Link').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(7, 40);
    reportSheet.getRange('H5').setValue('Stack').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.setColumnWidth(8, 150);
    reportSheet.getRange('I5').setValue('Extra stack').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
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
    let row = 6;
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

        if (trainingHours >= 10) {
            // Выделить строку зеленым цветом
            reportSheet.getRange(row, 2, 1, 11).setBackground("#d9ead3"); // Смените число 11 на число столбцов в вашей строке
        }

        reportSheet.getRange(row, 2).setValue(developerName).setVerticalAlignment("middle").setWrap(true);
        reportSheet.getRange(row, 3).setValue(developer.location).setVerticalAlignment("middle").setWrap(true);
        reportSheet.getRange(row, 4).setValue(englishLevel).setVerticalAlignment("middle");
        reportSheet.getRange(row, 5).setValue(trainingHours).setVerticalAlignment("middle");
        reportSheet.getRange(row, 6).setValue(salesHours).setVerticalAlignment("middle");
        let profileLink = getCompetenceData(developerName)['личное дело'] ?? '';
        if (profileLink) {
            reportSheet.getRange(row, 7).setFormula(`=HYPERLINK("${profileLink}", "Link")`).setVerticalAlignment("middle");
        }
        let competenceText = getCompetenceData(developerName)['Инструменты\nБиблиотеки\nСитстемы'] ?? '';
        reportSheet.getRange(row, 8).setValue(getCompetenceData(developerName)['Основной стек'] ?? '').setVerticalAlignment("middle");
        reportSheet.getRange(row, 9).setValue(getCompetenceData(developerName)['Дополнительный стек'] ?? '').setNote(competenceText).setVerticalAlignment("middle");
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
        let formula = `=SUMIF(XXX6:XXX${totalRow-1}, "<>", A6:A${totalRow-1})`;
        formula = formula.replace('XXX6:XXX', `${columnLetter}6:${columnLetter}`);
        formula = formula.replace('XXX6:A', `${columnLetter}6:A`);
        reportSheet.getRange(4, i).setFormula(formula);
    }


    if(all) {
        // The starting column for your formula for the remaining columns after the 3 empty ones
        let remainingStartColumn = endColumn + 3;  // Adding 3 to account for fact and plan columns and the column where the next set of data begins

        // Apply the second formula from remainingStartColumn to the last column
        let counter = 0;
        for(let i = remainingStartColumn; i <= lastColumn; i++) {
            let columnLetter = getColumnLetter(i);
            let formula = `=SUM(K6:K${totalRow-1})`;
            formula = formula.replace('K6:K', `${columnLetter}6:${columnLetter}`);

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

    // Add date and time of data gathering
    const currentTime = new Date().toLocaleString("en-GB", {timeZone: "Asia/Tbilisi"});
    reportSheet.getRange("B4").setValue(`Generated at ${currentTime} (Tbilisi, Georgia Timezone)`);

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




function getCompetencesOLD(sheet, developers) {
    const COMPETENCES_START_COLUMN = 6; // 'F' column
    const COMPETENCES_END_COLUMN = 55; // 'BD' column
    const DEVELOPERS_START_ROW = 3;
    const DEVELOPER_NAME_COLUMN = 3; // 'C' column
    const DEVELOPER_ENGLISH_LEVEL_COLUMN = 4; // 'D' column
    const DEVELOPER_JUNIOR_COLUMN = 5; // 'E' column

    const competenceDevelopers = {};
    let uniqueCompetences = [];
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    const data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    const competences = data[1].slice(COMPETENCES_START_COLUMN - 1, COMPETENCES_END_COLUMN);

    // Create a new array containing only developer names
    const developerNames = developers.map(developer => developer.name.split("(")[0].trim());

    for (let i = DEVELOPERS_START_ROW - 1; i < lastRow; i++) {
        const row = data[i];
        const developerName = row[DEVELOPER_NAME_COLUMN - 1];
        if (!developerName) break; // Stop loop when name is empty

        const developerEnglishLevel = row[DEVELOPER_ENGLISH_LEVEL_COLUMN - 1];
        const developerIsJunior = row[DEVELOPER_JUNIOR_COLUMN - 1] === 'да';
        const developerCompetences = row.slice(COMPETENCES_START_COLUMN - 1, COMPETENCES_END_COLUMN);

        competenceDevelopers[developerName] = {
            englishLevel: developerEnglishLevel,
            isJunior: developerIsJunior,
            competences: competences.reduce((acc, competence, index) => {
                acc[competence] = {
                    score: developerCompetences[index],
                    note: ''
                };
                return acc;
            }, {})
        };

        // Update the competences array only with competences of current developers that have scores
        if (developerNames.includes(developerName)) {
            let currentDeveloperCompetences = Object.entries(competenceDevelopers[developerName].competences);
            currentDeveloperCompetences.forEach(([key, value]) => {
                if (value.score && !uniqueCompetences.includes(key)) {
                    uniqueCompetences.push(key);
                }
            });
        }
    }

    return { competenceDevelopers, competences: uniqueCompetences };
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
    var startRow = 6; // Начальная строка, с которой нужно вставлять формулы
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


