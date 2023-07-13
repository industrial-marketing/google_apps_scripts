function saveOAuthToken() {
    var token = ScriptApp.getOAuthToken();
    PropertiesService.getScriptProperties().setProperty('OAUTH_TOKEN', token);

}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
        .addItem('Save OAuth Token', 'saveOAuthToken')
        .addItem('Update week plan from Workload 4 PM', 'updateWeekPlanCheckOAuth')
        .addItem('Update week fact from Week report sheet', 'updateWeekFact')
        .addItem('Update students total hours by stack', 'updateLearningTotals')
        .addItem('Update students total hours', 'calculateLearningTotals')
        .addItem('Update tickets total hours', 'updateTickets')
        .addItem('Generate weekly report', 'generateWeeklyReport')
        .addItem('Generate scrum report', 'generateScrumReport')
        .addItem('Show only learning rows', 'ShowOnlyLearnRows')
        .addItem('Show only HR rows', 'ShowOnlyHrRows')
        .addItem('Show all rows', 'showAllRows')
        .addToUi();
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

function updateWeekPlanCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию updateWeekPlan()
        updateWeekPlan();
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

function generateScrumReportCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию generateScrumReport()
        generateScrumReport();
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

function getHoursByNameAndProject(data, name) {
    var hoursAndProjects = [];
    for (var i = 0; i < data.length; i++) {
        var rowName = data[i][1].toString();
        //Logger.log(data[i][1]);
        if (rowName.startsWith(name)) {
            for (var j = 1; j < data[0].length; j++) {
                var cellValue = data[i][j];
                var hours = parseFloat(cellValue);
                if (hours > 0) {
                    var pm = data[0][j];
                    var project = data[4][j];
                    project = project.trim();
                    hoursAndProjects.push(pm + " " + project + " (" + hours.toFixed(2) + ")");
                }
            }
            break;
        }
    }
    return hoursAndProjects.join(', ');
}

function updateWeekPlan() {
    var studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    var sheet = studentsSheet;

    var sheetName = studentsSheet.getRange("G1").getValue();

    var externalSheetId = "1N65NUtqBA855C6K8swmeFQ9HbvIZU4fq4EnhYzvNV7Q"; // Замените на внешний ID таблицы
    var externalSpreadsheet = SpreadsheetApp.openById(externalSheetId);
    var externalSheet = externalSpreadsheet.getSheetByName(sheetName);

    var students = sheet.getRange("A2:A").getValues(); // Получаем значения из колонки A текущего листа
    var dataRange = externalSheet.getRange("C1:AN");
    var data = dataRange.getValues();
    var headers = externalSheet.getRange("5:5").getValues()[0]; // Получаем строки заголовков

    for (var i = 0; i < students.length; i++) {
        var studentName = students[i][0].toString();
        if (studentName == "") continue;
        Logger.log(studentName);
        var outputData = []; // массив для сбора данных
        var trainingHours = 0; // Счетчик часов обучения
        var hrHours = 0; // Счетчик HR часов
        var projectHours = getHoursByNameAndProject(data, studentName); // вызываем функцию, чтобы собрать информацию о проектах и часах
        Logger.log("HOURS " + JSON.stringify(projectHours));
        for (var j = 0; j < 100; j++) {
            for (var k = 0; k < data[j].length; k++) {
                if (data[j][1].toString().startsWith(studentName) && studentName != "") {
                    // Если это часы обучения или HR часы, то добавляем к соответствующему счетчику
                    if (headers[k+2] == "LEARN") {
                        trainingHours += data[j][k];
                        Logger.log("LEARN " + trainingHours);
                    }
                    else if (headers[k+2] == "HR") {
                        hrHours += data[j][k];
                        Logger.log("HR " + hrHours);
                    }
                }
            }
        }
        // Записываем часы обучения и HR часы в соответствующие колонки
        if(trainingHours > 0) sheet.getRange("G" + (i + 2)).setValue(trainingHours);
        else sheet.getRange("G" + (i + 2)).setValue("");
        if(hrHours > 0) sheet.getRange("J" + (i + 2)).setValue(hrHours);
        else sheet.getRange("J" + (i + 2)).setValue("");

        // Записываем собранные данные в колонку O
        if (projectHours != "") {
            sheet.getRange("M" + (i + 2)).setValue(projectHours);
        } else {
            sheet.getRange("M" + (i + 2)).setValue("");
        }
    }
}


function updateSheet(targetSheetName, targetColumn, additionalColumn) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var currentSheet = spreadsheet.getSheetByName("Students");
    var targetSheet = spreadsheet.getSheetByName(targetSheetName);
    var ticketsSheet = spreadsheet.getSheetByName("Tickets");

    var currentData = currentSheet.getDataRange().getValues();
    var targetData = targetSheet.getDataRange().getValues();
    var ticketsData = ticketsSheet.getDataRange().getValues();

    // Создаем словарь из данных таблицы "Tickets"
    var ticketsDict = {};
    for (var i = 0; i < ticketsData.length; i++) {
        ticketsDict[ticketsData[i][1]] = ticketsData[i][0];
    }

    var sumH = 0;
    var sumK = 0;
    var sum = 0;
    var foundValues = []; // массив для хранения найденных значений

    for (var i = 1; i < currentData.length; i++) {
        var currentName = currentData[i][1].toString(); // Значение из колонки B текущего листа
        if(currentName == "") continue;

        for (var j = 0; j < targetData.length; j++) {
            var targetName = targetData[j][0].toString(); // Значение из колонки A целевого листа
            var targetValue = targetData[j][2]; // Значение из колонки C целевого листа
            var targetKey = targetData[j][1].toString(); // Значение из колонки B целевого листа
            if (currentName === targetName) {
                if (targetKey.startsWith("LEARN")) {
                    sumH += targetValue;
                } else if (targetKey === "EXT-11" || targetKey === "EXT-91") {
                    sumK += targetValue;
                }
                sum += targetValue;

                var ticketValue = ticketsDict[targetKey]; // Получаем значение из словаря "Tickets" на основе targetData[j][1]
                if (ticketValue !== undefined) {
                    foundValues.push(ticketValue + " (" + targetValue.toFixed(2) + ")"); // добавляем значение из словаря "Tickets" и округленную targetValue в массив
                } else {
                    foundValues.push(targetKey + " (" + targetValue.toFixed(2) + ")"); // в случае если значения нет в словаре, добавляем само значение targetData[j][1] и округленную targetValue
                }
            }
        }

        if(sumH>0) currentSheet.getRange("H" + (i + 1)).setValue(sumH);
        else currentSheet.getRange("H" + (i + 1)).setValue("");

        if(sumK>0) currentSheet.getRange("K" + (i + 1)).setValue(sumK);
        else currentSheet.getRange("K" + (i + 1)).setValue("");

        if(additionalColumn) {
            if(foundValues.length > 0) currentSheet.getRange(additionalColumn + (i + 1)).setValue(foundValues.join(', ')); // заполняем колонку M найденными значениями
            else currentSheet.getRange(additionalColumn + (i + 1)).setValue("");
        }

        if(sum>0) currentSheet.getRange("P" + (i + 1)).setValue(sum);
        else currentSheet.getRange("P" + (i + 1)).setValue("");

        sumH = 0; // Сбрасываем сумму для следующей строки
        sumK = 0; // Сбрасываем сумму для следующей строки
        sum = 0; // Сбрасываем сумму для следующей строки
        foundValues = []; // очищаем массив найденных значений для следующей строки
    }
}

function updateWeekFact() {
    updateSheet("Week report All", "H", "N");
    updateSheet("Week report All", "K");
}

function updateLearningTotals() {
    var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    var learningTotalsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Learning totals");

    var currentData = currentSheet.getDataRange().getValues();
    var learningTotalsData = learningTotalsSheet.getDataRange().getValues();

    for (var i = 0; i < currentData.length; i++) {
        var currentCellValue = currentData[i][1];

        if (currentCellValue !== "") {
            //Logger.log("Ищем совпадения для значения " + currentCellValue + " в текущей строке " + (i + 1));

            for (var j = 0; j < learningTotalsData.length; j++) {
                var learningTotalsCellValue = learningTotalsData[j][0];

                if (currentCellValue === learningTotalsCellValue) {
                    //Logger.log("Теперь ищем совпадения для значения " + learningTotalsCellValue + " в текущей строке " + (j + 1));
                    var ticketCode = learningTotalsData[j][1];
                    //Logger.log(ticketCode);
                    var columnIndex = getMatchingColumnIndex(currentSheet, ticketCode);
                    //Logger.log(columnIndex);
                    var learningTotalsValue = learningTotalsData[j][2];
                    if (columnIndex > 2) {
                        //Logger.log("Найдено совпадение для значения " + currentCellValue + " на листе 'Learning totals'. Сумма: " + learningTotalsValue);
                        //var currentValue = currentData[i][columnIndex];
                        var newValue = learningTotalsValue;
                        currentSheet.getRange(i + 1, columnIndex + 1).setValue(newValue);
                        //Logger.log("Значение обновлено на текущем листе. Новое значение: " + newValue);
                    }
                }
            }
        }
    }
}

function getMatchingColumnIndex(sheet, value) {
    var firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    for (var i = 0; i < firstRow.length; i++) {
        if (firstRow[i] === value) {
            return i;
        }
    }

    return -1;
}


function updateTickets() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var ticketsSheet = spreadsheet.getSheetByName("Tickets");
    var learningTotalsSheet = spreadsheet.getSheetByName("Learning totals");

    var ticketsData = ticketsSheet.getDataRange().getValues();
    var learningTotalsData = learningTotalsSheet.getDataRange().getValues();

    // Обнуляем значения в колонке D на листе "Tickets"
    ticketsSheet.getRange("D2:D").clearContent();

    for (var i = 0; i < ticketsData.length; i++) {
        var ticketValue = ticketsData[i][1];

        if (ticketValue !== "") {
            var sum = 0;
            for (var j = 0; j < learningTotalsData.length; j++) {
                var learningTotalsValue = learningTotalsData[j][1];

                if (ticketValue === learningTotalsValue) {
                    var learningTotalsSum = learningTotalsData[j][2];
                    sum += learningTotalsSum;
                }
            }
            ticketsSheet.getRange(i + 1, 4).setValue(sum);
        }
    }
}


function calculateLearningTotals() {
    var studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    var learningTotalsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Learning totals");

    var studentsData = studentsSheet.getDataRange().getValues();
    var learningTotalsData = learningTotalsSheet.getDataRange().getValues();

    // Найти колонку, где в первой строке написано "Total"
    var totalColumnIndex = studentsData[0].indexOf("Total") + 1; // +1 потому что индексы в массивах начинаются с 0, а в Google Sheets - с 1

    // Очистить значения в колонке L на листе Students, только там, где в колонке A есть значение
    for (var i = 1; i < studentsData.length; i++) { // Начинаем с индекса 1, чтобы пропустить заголовки
        var studentName = studentsData[i][1]; // Значение из колонки B листа Students

        if (studentName !== "") {
            studentsSheet.getRange(i + 1, totalColumnIndex).clearContent();
        }
    }

    for (var i = 1; i < studentsData.length; i++) { // Начинаем с индекса 1, чтобы пропустить заголовки
        var studentName = studentsData[i][1]; // Значение из колонки B листа Students
        var totalHours = 0;

        if (studentName !== "") {
            for (var j = 1; j < learningTotalsData.length; j++) { // Начинаем с индекса 1, чтобы пропустить заголовки
                var learningTotalsName = learningTotalsData[j][0]; // Значение из колонки A листа Learning totals
                var hours = learningTotalsData[j][2]; // Значение из колонки C листа Learning totals

                if (studentName === learningTotalsName && hours !== "") {
                    totalHours += hours;
                }
            }

            // Записать сумму часов в колонку "Total" на листе Students
            studentsSheet.getRange(i + 1, totalColumnIndex).setValue(totalHours);
        }
    }
}


function ShowOnlyLearnRows() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    var data = sheet.getDataRange().getValues();

    sheet.showRows(1, data.length); // Обязательно показываем все строки перед скрытием

    for (var i = 0; i < data.length; i++) {
        // Проверяем колонки G, H, J, K (индексы 6, 7, 9, 10 соответственно)
        if (data[i][6] === "" && data[i][7] === "") {
            sheet.hideRows(i + 1);
        }
    }
    // Показать столбцы G, H, I
    sheet.showColumns(7, 3);

    // Скрыть столбцы J, K, L
    sheet.hideColumns(10, 3);
}

function ShowOnlyHrRows() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    var data = sheet.getDataRange().getValues();

    sheet.showRows(1, data.length); // Обязательно показываем все строки перед скрытием

    for (var i = 0; i < data.length; i++) {
        // Проверяем колонки G, H, J, K (индексы 6, 7, 9, 10 соответственно)
        if (data[i][9] === "" && data[i][10] === "") {
            sheet.hideRows(i + 1);
        }
    }
    // Показать столбцы J, K, L
    sheet.showColumns(10, 3);

    // Скрыть столбцы G, H, I
    sheet.hideColumns(7, 3);
}

function showAllRows() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    var data = sheet.getDataRange().getValues();

    sheet.showRows(1, data.length);
    // Показать все столбцы G, H, I, J, K, L
    sheet.showColumns(7, 6);
}

function generateWeeklyReport() {
    var studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    var reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reports");

    var studentsData = studentsSheet.getDataRange().getValues();
    studentsData = studentsData.filter(function(row) {
        return row[0] != "";
    });

    var reportTitle = "Learning report for " + studentsSheet.getRange("G1").getValue();

    var lastRow = reportSheet.getLastRow();
    var reportStartRow = lastRow > 0 ? lastRow + 5 : 1;

    // Создание жирной горизонтальной линии
    var lineRange = reportSheet.getRange(reportStartRow, 1, 1, reportSheet.getLastColumn());
    lineRange.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    lineRange.setFontWeight("bold");

    reportStartRow += 5;

    var cell = reportSheet.getRange("A" + reportStartRow);
    cell.setValue(reportTitle);
    var textStyle = {
        fontSize: 20,
        bold: true
    };
    cell.setFontSize(textStyle.fontSize);
    cell.setFontWeight(textStyle.bold);


    reportStartRow++;

    // Получаем заголовки из листа Students
    var headers = [studentsData[1][0], studentsData[1][3], studentsData[1][4], studentsData[1][5], studentsData[1][6], studentsData[1][7], studentsData[1][8], studentsData[1][12], studentsData[1][13], studentsData[1][14], studentsData[1][15]];  // Добавляем столбцы O и P

    // Незапланированные/превышенные часы по обучению
    var cell = reportSheet.getRange("A" + reportStartRow);
    cell.setValue("Незапланированные/превышенные часы по обучению");
    var textStyle = {
        fontSize: 14,
        bold: true
    };
    cell.setFontSize(textStyle.fontSize);
    cell.setFontWeight(textStyle.bold);

    reportSheet.getRange("A" + (reportStartRow + 1) + ":K" + (reportStartRow + 1)).setValues([headers]);  // Заголовки столбцов
    reportStartRow += 2;

    var excessHoursData = studentsData.filter(row => row[8] < 0).sort((a, b) => a[8] - b[8]); // Фильтрация и сортировка данных

    for (var i = 0; i < excessHoursData.length; i++) {
        var rowToInsert = [excessHoursData[i][0], excessHoursData[i][3], excessHoursData[i][4], excessHoursData[i][5], excessHoursData[i][6], excessHoursData[i][7], excessHoursData[i][8], excessHoursData[i][12], excessHoursData[i][13], excessHoursData[i][14], excessHoursData[i][15]];  // Добавляем столбцы O и P
        reportSheet.getRange("A" + (reportStartRow + i) + ":K" + (reportStartRow + i)).setValues([rowToInsert]);
    }

    reportStartRow += excessHoursData.length + 1;

    // Неотработанное/недоработанное обучение
    var cell = reportSheet.getRange("A" + reportStartRow);
    cell.setValue("Неотработанное/недоработанное обучение");
    var textStyle = {
        fontSize: 14,
        bold: true
    };
    cell.setFontSize(textStyle.fontSize);
    cell.setFontWeight(textStyle.bold);

    reportSheet.getRange("A" + (reportStartRow + 1) + ":K" + (reportStartRow + 1)).setValues([headers]);  // Заголовки столбцов
    reportStartRow += 2;

    var unworkedHoursData = studentsData.filter(row => row[8] > 0).sort((a, b) => a[8] - b[8]); // Фильтрация и сортировка данных

    for (var i = 0; i < unworkedHoursData.length; i++) {
        var rowToInsert = [unworkedHoursData[i][0], unworkedHoursData[i][3], unworkedHoursData[i][4], unworkedHoursData[i][5], unworkedHoursData[i][6], unworkedHoursData[i][7], unworkedHoursData[i][8], unworkedHoursData[i][12], unworkedHoursData[i][13], unworkedHoursData[i][14], unworkedHoursData[i][15]];  // Добавляем столбцы O и P
        reportSheet.getRange("A" + (reportStartRow + i) + ":K" + (reportStartRow + i)).setValues([rowToInsert]);
    }
}

function generateScrumReport() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var settingsSheet = spreadsheet.getSheetByName('Scrum report settings');
    var scrumFilesSheet = spreadsheet.getSheetByName('Scrum files');
    var reportSheet = spreadsheet.getSheetByName('Scrum report');
    var studentsSheet = spreadsheet.getSheetByName('Students');
    var jiraReportSheet = spreadsheet.getSheetByName('Jira report All');
    var replaceSheet = spreadsheet.getSheetByName('Scrum report project name replace');

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

    var scrumFilesData = scrumFilesSheet.getRange(3, 1, scrumFilesSheet.getLastRow(), 3).getValues();
    var studentsData = studentsSheet.getDataRange().getValues();
    var settingsData = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, settingsSheet.getLastColumn()).getValues();
    var replaceData = replaceSheet.getRange(1, 1, replaceSheet.getLastRow(), 2).getValues();
    var jiraReportData = jiraReportSheet.getRange(3, 1, jiraReportSheet.getLastRow(), 7).getValues();

    // Удаление строк с null значением в первой колонке или недопустимой датой
    jiraReportData = jiraReportData.filter(function(row) {
        var dateValue = Utilities.formatDate(new Date(row[0]), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
        return dateValue !== "null" && dateValue !== "01/01/1970";
    });

    // Проверка, что есть хотя бы одна строка после фильтрации
    if (jiraReportData.length === 0) {
        console.warn("Warning: No valid data after filtering");
        return;
    }

    // Проверка, что все значения в первой колонке имеют один и тот же месяц
    var firstMonth = Utilities.formatDate(new Date(jiraReportData[0][0]), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "MM");
    for (var i = 1; i < jiraReportData.length; i++) {
        var currentMonth = Utilities.formatDate(new Date(jiraReportData[i][0]), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "MM");
        if (currentMonth !== firstMonth) {
            // Вывод предупреждения и прерывание выполнения функции
            console.warn("Warning: Values in the first column do not have the same month because " + currentMonth + " is not " + firstMonth);
            return;
        }
    }

    // Проверка, что даты идут по порядку
    for (var i = 1; i < jiraReportData.length; i++) {
        var currentDate = new Date(jiraReportData[i][0]);
        var previousDate = new Date(jiraReportData[i-1][0]);
        if (currentDate < previousDate) {
            // Вывод предупреждения и прерывание выполнения функции
            console.warn("Warning: Dates in the first column are not in ascending order");
            return;
        }
    }


    function findTypeAndProject(ticketNumber, sprintName, projectName, summary) {
        for (var i = 0; i < settingsData.length; i++) {
            var ticketCondition = settingsData[i][0];
            var sprintCondition = settingsData[i][1];
            var summaryCondition = settingsData[i][2];
            var notSummaryCondition = settingsData[i][3];
            var typeValue = settingsData[i][4];
            var projectValue = settingsData[i][5];

            var ticketMatches = false;
            var sprintMatches = false;
            var summaryMatches = false;
            var notSummaryMatches = false;

            if (projectValue === '') {
                if (!ticketNumber.toLowerCase().startsWith('ext-')) {
                    projectValue = projectName;
                } else {
                    projectValue = sprintName;
                }
            }

            if (typeValue === '') {
                typeValue = 'DEV';
            }

            // Check ticket condition
            if (ticketCondition) {
                var ticketConditions = ticketCondition.split('|');
                for (var j = 0; j < ticketConditions.length; j++) {
                    var condition = ticketConditions[j].trim();
                    if (condition.startsWith('*') && condition.endsWith('*')) {
                        var keyword = condition.substring(1, condition.length - 1);
                        if (ticketNumber.toLowerCase().includes(keyword)) {
                            ticketMatches = true;
                            break;
                        }
                    } else if (condition.startsWith('*')) {
                        var keyword = condition.substring(1);
                        if (ticketNumber.toLowerCase().endsWith(keyword)) {
                            ticketMatches = true;
                            break;
                        }
                    } else if (condition.endsWith('*')) {
                        var keyword = condition.substring(0, condition.length - 1);
                        if (ticketNumber.toLowerCase().startsWith(keyword)) {
                            ticketMatches = true;
                            break;
                        }
                    } else if (ticketNumber.toLowerCase() == condition.toLowerCase()) {
                        ticketMatches = true;
                        break;
                    }
                }
            } else {
                // If no ticket condition is specified, consider it a match
                ticketMatches = true;
            }

            // Check sprint condition
            if (sprintCondition) {
                var sprintConditions = sprintCondition.split('|');
                for (var j = 0; j < sprintConditions.length; j++) {
                    var condition = sprintConditions[j].trim();
                    if (condition.startsWith('*') && condition.endsWith('*')) {
                        var keyword = condition.substring(1, condition.length - 1);
                        if (sprintName.toLowerCase().includes(keyword)) {
                            sprintMatches = true;
                            break;
                        }
                    } else if (condition.startsWith('*')) {
                        var keyword = condition.substring(1);
                        if (sprintName.toLowerCase().endsWith(keyword)) {
                            sprintMatches = true;
                            break;
                        }
                    } else if (condition.endsWith('*')) {
                        var keyword = condition.substring(0, condition.length - 1);
                        if (sprintName.toLowerCase().startsWith(keyword)) {
                            sprintMatches = true;
                            break;
                        }
                    } else if (sprintName.toLowerCase() == condition.toLowerCase()) {
                        sprintMatches = true;
                        break;
                    }
                }
            } else {
                // If no sprint condition is specified, consider it a match
                sprintMatches = true;
            }

            // Check summary condition
            if (summaryCondition) {
                var summaryConditions = summaryCondition.split('|');
                for (var j = 0; j < summaryConditions.length; j++) {
                    var condition = summaryConditions[j].trim();
                    if (condition.startsWith('*') && condition.endsWith('*')) {
                        var keyword = condition.substring(1, condition.length - 1);
                        if (summary.toLowerCase().includes(keyword)) {
                            summaryMatches = true;
                            break;
                        }
                    } else if (condition.startsWith('*')) {
                        var keyword = condition.substring(1);
                        if (summary.toLowerCase().endsWith(keyword)) {
                            summaryMatches = true;
                            break;
                        }
                    } else if (condition.endsWith('*')) {
                        var keyword = condition.substring(0, condition.length - 1);
                        if (summary.toLowerCase().startsWith(keyword)) {
                            summaryMatches = true;
                            break;
                        }
                    } else if (summary.toLowerCase() == condition.toLowerCase()) {
                        summaryMatches = true;
                        break;
                    }
                }
            } else {
                // If no summary condition is specified, consider it a match
                summaryMatches = true;
            }

            // Check not summary condition
            if (notSummaryCondition) {
                if (!summary.toLowerCase().includes(notSummaryCondition.toLowerCase())) {
                    notSummaryMatches = true;
                }
            } else {
                // If no not summary condition is specified, consider it a match
                notSummaryMatches = true;
            }

            // If all conditions match, set type and project
            if (ticketMatches && sprintMatches && summaryMatches && notSummaryMatches) {
                return {
                    type: typeValue,
                    project: replaceProjectName(projectValue)
                };
            }
        }

        // If no conditions match, return default values
        return {
            type: 'ERROR:' + ticketNumber + ' ' + summary,
            project: 'ERROR:' + sprintName
        };
    }


    // Function to replace project name using "Scrum report project name replace" sheet
    function replaceProjectName(projectName) {
        for (var i = 0; i < replaceData.length; i++) {
            var currentValue = replaceData[i][0];
            var replaceValue = replaceData[i][1];
            if (projectName === currentValue) {
                return replaceValue;
            }
        }
        return projectName;
    }


    var reportColumn = 1;
    var startRow = reportSheet.getLastRow() + 5;

    for (var i in scrumFilesData) {
        var name = scrumFilesData[i][0];
        var fileLink = scrumFilesData[i][1];
        var pm = scrumFilesData[i][2];

        // If fileLink is empty or undefined, skip the current iteration
        if (!fileLink) {
            console.warn('No file link found for name: ' + name);
            continue;
        }

        var correspondingStudent = studentsData.find(function(row) {
            if (typeof row[0] === 'string' && row[0].startsWith(name)) {
                return row[0].startsWith(name);
            }
            return false;
        });

        if (!correspondingStudent) {
            console.warn('No corresponding developer found for name: ' + name);
            continue;
        }

        var englishName = correspondingStudent[1];

        var groupData = {};
        var groupCommentsData = {};
        if(jiraReportData.length == 0) {
            console.warn('No jiraReportData found for name: ' + name);
            continue;
        }
        jiraReportData.forEach(function(row) {
            if (row[1] === englishName) {
                var date = Utilities.formatDate(new Date(row[0]), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
                var projectName = row[2];
                var ticketNumber = row[3];
                var summary = row[4];
                var hours = row[6];
                var sprintName = row[5];
                var comments = row[3] + " - " + row[4] + " - " + row[6] + "                                                        ";

                var typeAndProject = findTypeAndProject(ticketNumber, sprintName, projectName, summary);
                var type = typeAndProject.type;
                var project = typeAndProject.project;

                if (pm == 'pm' && type == 'DEV') type = 'PM';
                if (pm == 'pm' && type == 'DEVfree') type = 'PMfree';
                if (pm == 'pm' && type == 'DEVint') type = 'PMint';

                var key = date + '|||' + type + '|||' + project;
                if (!groupData[key]) {
                    groupData[key] = hours;
                    groupCommentsData[key] = comments;
                } else {
                    groupData[key] += hours;
                    groupCommentsData[key] += comments;
                }
            }
        });


        var reportData = [];
        // var scrumData = [];
        var hoursTotal = 0;
        for (var key in groupData) {
            var [date, type, project] = key.split('|||');
            var hours = groupData[key];
            var comments = groupCommentsData[key];
            hoursTotal += hours;
            reportData.push([name, date, type, project, hours, comments]);
        }

        var month = new Date(date.split('/')[2], date.split('/')[1] - 1, date.split('/')[0]).toLocaleString('ru-RU', { month: 'long' }).charAt(0).toUpperCase() + new Date(date.split('/')[2], date.split('/')[1] - 1, date.split('/')[0]).toLocaleString('ru-RU', { month: 'long' }).slice(1);
        var importRangeFormula = '=QUERY(IMPORTRANGE("' + fileLink + '", "' + month + '!A2:E900"), "SELECT Col1, Col2, Col3, Col5")';
        var columnLetter = getColumnLetter(reportColumn + 8);
        var hoursScrumTotalFormula = '=SUM(' + columnLetter + '8:' + columnLetter + '909)';

        // старый вариант с использованием сервиса Sheets (убрали его, чтобы не было оплаты за API)
        // var file = SpreadsheetApp.openByUrl(fileLink);
        // var monthSheet = file.getSheetByName(month);
        // var monthSheetData = monthSheet.getRange(2, 1, monthSheet.getLastRow() - 1, 5).getValues();
        // var hoursScrumTotal = 0;
        // if(monthSheetData.length == 0) {
        //   console.warn('No monthSheetData found for name: ' + name);
        //   continue;
        // }
        // monthSheetData.forEach(function(rowData) {
        //   if (rowData[0] && rowData[1] && rowData[2] && rowData[4]) {
        //     var dateScrum = Utilities.formatDate(rowData[0], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
        //     var typeScrum = rowData[1];
        //     var projectScrum = rowData[2];
        //     var hoursScrum = rowData[4];
        //     hoursScrumTotal += hoursScrum;
        //     scrumData.push([name, dateScrum, typeScrum, projectScrum, hoursScrum]);
        //   }
        // });

        var row = startRow;

        reportSheet.getRange(row++, reportColumn).setValue(name);
        reportSheet.getRange(row++, reportColumn).setValue(fileLink);
        reportSheet.getRange(row++, reportColumn, 1, 5).setValues([[ 'Jira', '', '', '', hoursTotal ]]);

        reportData.forEach(function(rowData) {
            var range = reportSheet.getRange(row++, reportColumn, 1, 5); // Указываем диапазон только для пяти столбцов
            range.setValues([rowData.slice(0, 5)]); // Используем только первые пять элементов массива rowData

            var commentText = rowData[5]; // Получаем текст комментария из последнего элемента массива rowData (comments)
            var cell = range.getCell(1, 5); // Получаем ячейку с часами (пятая колонка)
            cell.setNote(commentText); // создаем новую заметку (комментарий) с содержанием из commentText
        });

        var encodedName = encodeURIComponent(name);
        var encodedFileLink = encodeURIComponent(fileLink);
        var hyperlinkUrl = "https://script.google.com/a/macros/sharp-dev.net/s/AKfycby35CX3LS2j9dfeISTH7HrEPNGhX7v2AWNx3wYpK9USYbhfZK2gsgAmh9Vr7wa0WE8z/exec?name=" + encodedName + "&url=" + encodedFileLink;
        reportSheet.getRange(row++, reportColumn).setFormula('=HYPERLINK("' + hyperlinkUrl + '", "Update Scrum file with these values above")');

        reportColumn += 5;

        var row = startRow;

        reportSheet.getRange(row++, reportColumn).setValue('');
        reportSheet.getRange(row++, reportColumn).setValue('');
        reportSheet.getRange(row, reportColumn).setValue('Scrumfile');
        reportSheet.getRange(row, reportColumn + 3).setFormula(hoursScrumTotalFormula);
        row++;

        // старый вариант
        // reportSheet.getRange(row++, reportColumn + 3).setFormula(hoursScrumTotalFormula);
        // reportSheet.getRange(row++, reportColumn, 1, 4).setValues([[ 'Scrumfile', '', '', hoursScrumTotalFormula ]]);

        reportSheet.getRange(row++, reportColumn).setFormula(importRangeFormula);

        // старый вариант
        // scrumData.forEach(function(rowData) {
        //   reportSheet.getRange(row++, reportColumn, 1, rowData.length).setValues([rowData]);
        // });

        // для старого API варианта было 5 колонок, включая повторяющееся имя в первой колонке для визуальной прозрачности, но если потребуется вернуть API, можно и на 4 перейти
        // reportColumn += 5;
        reportColumn += 4;
        var lastCell = reportSheet.getRange(row - 1, reportColumn);
        var lastRow = reportSheet.getLastRow();

        var range = reportSheet.getRange(1, lastCell.getColumn()-1, lastRow, 1);
        range.setBorder(null, null, null, true, null, null, 'bold', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    }
}

function getColumnLetter(column) {
    var letter = "";
    while (column > 0) {
        var remainder = (column - 1) % 26;
        letter = String.fromCharCode(65 + remainder) + letter;
        column = Math.floor((column - 1) / 26);
    }
    return letter;
}

