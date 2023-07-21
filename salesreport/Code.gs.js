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
        .addItem('Обновить "Scrum files for current week"', 'gatherDataInCurrentSheetCheckOAuth')
        .addItem('Обновить "Competences"', 'copyDataToCompetencesSheetCheckOAuth')
        .addItem('Обновить "DeveloperStackData"', 'updateDeveloperStackDataCheckOAuth')
        .addToUi();
    ui.createMenu('Фильтры')
        .addItem('Показать всех', 'showAllRows')
        .addItem('Только бенч', 'showOnlyBenchRows')
        .addItem('Выбор стеков', 'showStacksDialog')
        //.addItem('Выбор проектов', 'showProjectsDialog')
        .addItem('Сортировать по A-Z', 'sortDataAscending')
        .addItem('Сортировать по Z-A', 'sortDataDescending')
        .addToUi();
}

function showStacksDialog() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Stacks.html')
        .setWidth(400)
        .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Стеки');
}

function showProjectsDialog() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Projects.html')
        .setWidth(400)
        .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Проекты');
}

function sortData(stackName) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report" или "SALES report"');
        return;
    }

    var headerRow = sheet.getRange("5:5"); // Это строка с именами стеков
    var values = headerRow.getValues();
    var sortColumn = values[0].indexOf(stackName) + 1; // Находим номер столбца для этого стека

    if (sortColumn > 0) { // Если стек найден в заголовке
        var dataRange = sheet.getRange(6, 1, sheet.getLastRow() - 6, sheet.getLastColumn());
        showAllRows();
        dataRange.sort([{column: sortColumn, ascending: true}]);
        hideEmptyRows(sheet, sortColumn);
    }
    else {
        Logger.log("Не удалось найти стек " + stackName);
    }
}

function filterByProject(projectName) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report" или "SALES report"');
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
}

function sortDataAscending() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report" или "SALES report"');
        return;
    }

    var sortColumn = sheet.getActiveCell().getColumn();

    var dataRange = sheet.getRange(6, 1, sheet.getLastRow() - 5, sheet.getLastColumn());

    showAllRows();
    dataRange.sort([{column: sortColumn, ascending: true}]);
}

function sortDataDescending() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report" или "SALES report"');
        return;
    }

    var sortColumn = sheet.getActiveCell().getColumn();

    var dataRange = sheet.getRange(6, 1, sheet.getLastRow() - 5, sheet.getLastColumn());

    showAllRows();
    dataRange.sort([{column: sortColumn, ascending: false}]);
}


function hideEmptyRows(sheet, sortColumn) {
    var sheetName = sheet.getName();

    // Проверяем, запущена ли функция на правильном листе
    if (sheetName !== 'ALL report' && sheetName !== 'SALES report') {
        Logger.log('Эта функция может быть запущена только на листах "ALL report" или "SALES report"');
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


function generateSalesReportUpdate() {
    generateSalesReport(false);
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


function generateSalesReportCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию generateSalesReport()
        generateSalesReport(false);
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


function gatherDataInCurrentSheetCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию updateWeekPlan()
        gatherDataInCurrentSheet();
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


function generateSalesReport(all = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let reportName = "SALES report"
    if(all) {
        reportName = 'ALL report';
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
    reportSheet.clear();

    const workloadSheetId = "1N65NUtqBA855C6K8swmeFQ9HbvIZU4fq4EnhYzvNV7Q";
    const workloadSpreadsheet = SpreadsheetApp.openById(workloadSheetId);

    const currentWeekMondayDate = new Date();
    currentWeekMondayDate.setDate(currentWeekMondayDate.getDate() - currentWeekMondayDate.getDay() + (currentWeekMondayDate.getDay() == 0 ? -6:1));
    const currentWeekSundayDate = new Date();
    currentWeekSundayDate.setDate(currentWeekSundayDate.getDate() - currentWeekSundayDate.getDay() + 7);

    const currentWeekMondayString = Utilities.formatDate(currentWeekMondayDate, ss.getSpreadsheetTimeZone(), 'd MMM').toLowerCase();
    const currentWeekSundayString = Utilities.formatDate(currentWeekSundayDate, ss.getSpreadsheetTimeZone(), 'd MMM').toLowerCase();

    let workloadSheetName = currentWeekMondayDate.getMonth() === currentWeekSundayDate.getMonth() ?
        `${currentWeekMondayString.split(" ")[0]}-${currentWeekSundayString.split(" ")[0]} ${currentWeekSundayString.split(" ")[1]}` :
        `${currentWeekMondayString}-${currentWeekSundayString.split(" ")[0]}`;

    const workloadSheet = workloadSpreadsheet.getSheetByName(workloadSheetName);
    if (!workloadSheet) {
        SpreadsheetApp.getUi().alert(`Cannot find sheet "${workloadSheetName}" in the workload spreadsheet.`);
        return;
    }


    let developers = getDevelopers(workloadSheet, all);

    Logger.log(developers.length);

    showAllRows();

    // Initialize report
    reportSheet.clearContents();
    reportSheet.getRange('B3').setValue( reportName + ` for ${currentWeekMondayString} - ${currentWeekSundayString}`).setFontSize(20);
    reportSheet.getRange('L3').setValue('для сортировки выделите ячейку в 5й строке и нажмите "Сортировать" или используйте дополнительные инструменты поиска меню "Фильтры"').setFontSize(9);

    // Initialize the header row
    reportSheet.getRange('B5').setValue('Developer').setVerticalAlignment("middle");
    reportSheet.getRange('C5').setValue('English').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('D5').setValue('Training').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('E5').setValue('Sales').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('F5').setValue('Plan').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('G5').setValue('Fact').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('H5').setValue('Profile Link').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('I5').setValue('Stack').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('J5').setValue('Extra stack').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('K5').setValue('Обучаемость\nСтрессоустойчивость\nРабота в команде\nРабота с клиентом\nНавыки самопрезентации\nГибкость мышления').setTextRotation(90).setBackground("#ffffff").setHorizontalAlignment("center").setVerticalAlignment("middle");

    let column = 12;
    let allStacks = {};

    for (let developer of developers) {
        if(!developer.name) continue;
        let developerName = developer.name.split("(")[0].trim();
        let stackData = getDeveloperStackDataFromSheet(developerName);

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


    // Initialize the data rows
    let row = 6;
    for (let developer of developers) {
        if(!developer.name) continue;
        let developerName = developer.name.split("(")[0].trim(); // Remove everything after the "(" and trim spaces
        let competenceData = getDeveloperCompetenceData(developerName)
        let englishLevel = competenceData['Английский'];
        // Here you need to calculate trainingAndSales and allocation for each developer
        let trainingHours = developer.projects['Training'] || 0;
        let salesHours = developer.projects['SALES'] || 0;

        let stackData = getDeveloperStackDataFromSheet(developerName);

        if (trainingHours >= 10) {
            // Выделить строку зеленым цветом
            reportSheet.getRange(row, 2, 1, 9).setBackground("#d9ead3"); // Смените число 10 на число столбцов в вашей строке
        }

        reportSheet.getRange(row, 2).setValue(developerName).setVerticalAlignment("middle");
        reportSheet.getRange(row, 3).setValue(englishLevel).setVerticalAlignment("middle");

        reportSheet.getRange(row, 4).setValue(trainingHours).setVerticalAlignment("middle");
        reportSheet.getRange(row, 5).setValue(salesHours).setVerticalAlignment("middle");
        reportSheet.getRange(row, 6).setValue(developer.projectHours).setVerticalAlignment("middle").setWrap(true);
        reportSheet.getRange(row, 7).setValue(getAllocationData(developers, developerName)).setVerticalAlignment("middle").setWrap(true);
        let profileLink = competenceData['личное дело'] ?? '';
        if (profileLink) {
            reportSheet.getRange(row, 8).setFormula(`=HYPERLINK("${profileLink}", "Link")`).setVerticalAlignment("middle");
        }
        let competenceText = competenceData['Инструменты\nБиблиотеки\nСитстемы'] ?? '';
        reportSheet.getRange(row, 9).setValue(competenceData['Основной стек'] ?? '').setVerticalAlignment("middle");
        reportSheet.getRange(row, 10).setValue(competenceData['Дополнительный стек'] ?? '').setNote(competenceText).setVerticalAlignment("middle");
        reportSheet.getRange(row, 11).setValue(
            (competenceData['Обучаемость'] ?? '') + '  ' +
            (competenceData['Стрессоустойчивость'] ?? '') + '  ' +
            (competenceData['Работа в команде'] ?? '') + '  ' +
            (competenceData['Работа с клиентом (командой клиента)'] ?? '') + '  ' +
            (competenceData['Навыки самопрезентации'] ?? '') + '  ' +
            (competenceData['Гибкость мышления'] ?? '')
        ).setVerticalAlignment("middle").setHorizontalAlignment("center");


        let column = 12;
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
        row++;
    }

    // Set the border
    reportSheet.getRange(5, 2, row-5, column-2).setBorder(true, true, true, true, true, true);
    insertSumFormulas(all);
    let lastColumn = reportSheet.getLastColumn();

    // определите номер строки, куда нужно вставить итоговые значения (после последней строки с данными)
    let totalRow = reportSheet.getLastRow() + 1;

    reportSheet.getRange(4, 11).setValue('доступные ресурсы:');
    // начиная со столбца L (12 в системе A1) и до последнего столбца
    for(let i = 12; i <= lastColumn; i++) {
        // вы можете использовать функцию SUMIF в Google Sheets, которая будет суммировать значения в столбце A
        // для тех строк, где значение в данном столбце (i) непустое
        let formula = `=SUMIF(L6:L${totalRow-1}, "<>", A6:A${totalRow-1})`;

        // заменим L на текущий столбец в цикле
        formula = formula.replace(/L/g, getColumnLetter(i));

        // вставляем формулу в ячейку
        reportSheet.getRange(4, i).setFormula(formula);
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


function getAllocationData(developers, developerName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scrumSheet = ss.getSheetByName('Scrum files for current week');

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
        let developerFull = developers.find(developer => developer.name.startsWith(developerShort));
        if (developerFull && developerFull.name) {
            developerFull.name = developerFull.name.split("(")[0].trim();
            if (!allocationData[developerFull.name]) {
                allocationData[developerFull.name] = {};
            }
            if (!allocationData[developerFull.name][project]) {
                allocationData[developerFull.name][project] = 0;
            }
            allocationData[developerFull.name][project] += hours;
            Logger.log(developerFull.name + ' ' + project + ' ' + hours);
        }
    });

    let allocationList = [];
    let totalDeveloperHours = 0;

    developerName = developerName.split("(")[0].trim();

    for (let project in allocationData[developerName]) {
        let roundedHours = Math.round(allocationData[developerName][project] * 100) / 100;
        totalDeveloperHours += roundedHours;
        allocationList.push(project + ' (' + roundedHours + ')');
        Logger.log(project + ' (' + roundedHours + ')');
    }

    allocationList.unshift(Math.round(totalDeveloperHours * 100) / 100 + '');
    allocationData[developerName] = allocationList.join(' | ');

    return allocationData[developerName];
}


function getDevelopers(workloadSheet, all) {
    let developers = [];
    let projects = [];

    let workloadData = workloadSheet.getDataRange().getValues();

    // Retrieve projects from the 5th row
    projects = workloadData[4].slice(1);

    // Iterate through the rows of the workloadData
    for (let i = 5; i < workloadData.length; i++) {
        // Get the developer's name, which is assumed to be in the 4th column
        let developerName = workloadData[i][3];
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
        let developer = {name: developerName, projectHours, projects: {}};
        let workedOnTraining = false;
        let workedOnSales = false;

        for (let j = 5; j < workloadData[i].length; j++) {
            hours = workloadData[i][j] || 0;
            let projectName = projects[j - 1];

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

function getCompetences(sheet, developers) {
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


function getScrumFilesDataOld(lastWeekMondayDate, lastWeekSundayDate) {
    const scrumFilesFolderId = "1AnMMx9rnnQE7r2KoodgYZP1eukycY0lJ";
    const folder = DriveApp.getFolderById(scrumFilesFolderId);
    const monthNames = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
    Logger.log(lastWeekMondayDate + ' ' + lastWeekSundayDate);
    // Create an array for month names between lastWeekMondayDate and lastWeekSundayDate
    let monthsToProcess = [];
    for (let date = new Date(lastWeekMondayDate); date <= lastWeekSundayDate; date.setMonth(date.getMonth() + 1)) {
        monthsToProcess.push(monthNames[date.getMonth()]);
    }

    Logger.log(`Looking for files: ${monthsToProcess.join(', ')}`);

    const filesIterator = folder.getFiles();
    let scrumFiles = [];
    while (filesIterator.hasNext()) {
        const file = filesIterator.next();
        Logger.log(`Checking file: ${file.getName()}`);
        if (monthsToProcess.some(monthName => file.getName().includes(monthName))) {
            scrumFiles.push(file);
        }
    }

    if(scrumFiles.length == 0) {
        Logger.log('No matching files found.');
        return;
    }

    let data = {};
    for (let i = 0; i < scrumFiles.length; i++) {
        let file = scrumFiles[i];
        Logger.log(`Processing file: ${file.getName()}`);
        const fileId = file.getId();
        const ss = SpreadsheetApp.openById(fileId);

        // Iterate over all required months for each file
        for (let monthName of monthsToProcess) {
            const sheet = ss.getSheetByName('СкрамФайлы');

            if (sheet) {
                let columnNumber = 2;
                while (true) {
                    const columnLetter = getColumnLetter(columnNumber);
                    const urlCell = sheet.getRange(`${columnLetter}4`);
                    const url = urlCell.getValue();
                    if (url === "") {
                        Logger.log(`No more URLs found in column ${columnLetter}.`);
                        break;
                    }
                    Logger.log(`Opening URL: ${url}`);
                    const externalFile = SpreadsheetApp.openByUrl(url);
                    const externalSheet = externalFile.getSheetByName(monthName);
                    const lastRow = externalSheet.getLastRow();
                    Logger.log(`Reading data from URL: ${url}, ${lastRow} rows found.`);

                    const monthSheetData = externalSheet.getRange(2, 1, lastRow - 1, 5).getValues();

                    if(monthSheetData.length == 0) {
                        Logger.log('No monthSheetData found.');
                        break;
                    }

                    monthSheetData.forEach(function(rowData) {
                        if (rowData[0] && rowData[1] && rowData[2] && rowData[4]) {
                            const dateTime = new Date(rowData[0]);
                            if (dateTime >= lastWeekMondayDate && dateTime <= lastWeekSundayDate) {
                                const developer = externalFile.getName(); // Here we reference the filename of the external file
                                const dateScrum = Utilities.formatDate(rowData[0], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
                                const typeScrum = rowData[1];
                                const projectScrum = rowData[2];
                                const hoursScrum = rowData[4];
                                if (!data[developer]) {
                                    data[developer] = [];
                                }
                                data[developer].push({date: dateScrum, type: typeScrum, project: projectScrum, hours: hoursScrum});
                                Logger.log(`${developer}, ${dateScrum}, ${typeScrum}, ${projectScrum}, ${hoursScrum}`);
                            }
                        }
                    });

                    columnNumber += 6;
                }
            } else {
                Logger.log(`Sheet "${monthName}" not found in file "${file.getName()}"`);
            }
        }
    }

    Logger.log(JSON.stringify(data)); // Log the entire data object
    return data;
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


function getCurrentMonday() {
    const today = new Date();
    const day = today.getDay();
    const diffToMonday = day === 0 ? 6 : day - 1; // If today is Sunday (0), we need to subtract 6 to get to the last Monday. Otherwise, subtract the number of days up to Monday
    const currentMonday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - diffToMonday);
    return currentMonday;
}


function getCurrentSunday() {
    const today = new Date();
    const day = today.getDay();
    const diffToNextSunday = (7 - day) % 7; // Here we calculate the number of days remaining to next Sunday
    const currentSunday = new Date(today.getFullYear(), today.getMonth(), today.getDate() + diffToNextSunday);
    return currentSunday;
}


function gatherDataInCurrentSheet() {
    const currentWeekMondayDate = getCurrentMonday();
    const currentWeekSundayDate = getCurrentSunday();

    const data = getScrumFilesData(currentWeekMondayDate, currentWeekSundayDate);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let dataSheet = ss.getSheetByName("Scrum files for current week");

    // Создаем новый лист, если он не существует
    if (!dataSheet) {
        dataSheet = ss.insertSheet("Scrum files for current week");
    } else {
        // Очистить лист, если он уже существует
        dataSheet.clear();
    }

    // Добавление даты и времени сбора данных
    const currentTime = new Date().toLocaleString("en-GB", {timeZone: "Asia/Tbilisi"});
    dataSheet.getRange("A1").setValue(`Data gathered at ${currentTime} (Tbilisi, Georgia Timezone)`);

    // Заголовки для нового листа смещены на строку вниз
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
    competencesSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);

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
    const headers = competencesData[0]; // First row contains the headers
    let developerCompetenceData = {};

    // Loop through the rest of the rows to find the matching developer
    for (let i = 1; i < competencesData.length; i++) {
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


function showOnlyBenchRows() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SALES report");
    var data = sheet.getDataRange().getValues();

    sheet.showRows(1, data.length); // Обязательно показываем все строки перед скрытием

    for (var i = 5; i < data.length; i++) {
        // Проверяем колонки D (индекс 3 соответственно)
        if (data[i][3] < 10) {
            sheet.hideRows(i + 1);
        }
    }
}


function showAllRows() {
    //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SALES report");
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    sheet.showRows(1, data.length);
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

    var currentTime = new Date();
    currentTime.setTime(currentTime.getTime() + 4 * 60 * 60 * 1000);
    sheet.appendRow([currentTime.toString()]);
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

function insertSumFormulas(all) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let sheetName = "SALES report"
    if(all) {
        sheetName = 'ALL report';
    }

    sheet = ss.getSheetByName(sheetName);

    var lastRow = sheet.getLastRow();
    var startRow = 6; // Начальная строка, с которой нужно вставлять формулы
    var startColumn = 1; // Начальный столбец, для которого нужно вставить формулы
    var endColumn = 1; // Последний столбец, для которого нужно вставить формулы

    for (var row = startRow; row <= lastRow; row++) {
        for (var column = startColumn; column <= endColumn; column++) {
            var cell = sheet.getRange(row, column);
            var formula = "=SUM(D" + row + ":E" + row + ")";
            cell.setFormula(formula).setVerticalAlignment("middle");
        }
    }
}


function getStacks() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    if (sheetName !== 'ALL report' && sheetName !== 'SALES report') {
        throw 'Error: This feature is only available for the "ALL report" or "SALES report" sheets.';
    }

    var stackRow = sheet.getRange("5:5");
    var values = stackRow.getValues();
    var stacks = values[0].slice(11); // Получить все значения начиная с колонки L
    return stacks;
}
