function saveOAuthToken() {
    var token = ScriptApp.getOAuthToken();
    PropertiesService.getScriptProperties().setProperty('OAUTH_TOKEN', token);

}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
        .addItem('Save OAuth Token', 'saveOAuthToken')
        .addItem('Update week plan/fact from Sales Report', 'updateStudentDataFromSalesReportCheckOAuth')
        .addItem('Update week plan from Workload 4PM', 'updateWeekPlanCheckOAuth')
        .addItem('Update week fact from Week report sheet', 'updateWeekFact')
        .addItem('Update students total hours by stack', 'updateLearningTotals')
        .addItem('Update students total hours', 'calculateLearningTotals')
        .addItem('Update tickets total hours', 'updateTickets')
        .addItem('Update students stacks', 'updateStudentStackCheckOAuth')
        .addItem('Generate weekly report', 'generateWeeklyReport')
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

function updateStudentStackCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию updateStudentStack()
        updateStudentStack();
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

function updateStudentDataFromSalesReportCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию updateStudentDataFromSalesReport()
        updateStudentDataFromSalesReport();
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
    var dataRange = externalSheet.getRange("C1:BB100");
    // let workloadData = workloadSheet.getDataRange().getValues();
    var data = dataRange.getValues();
    var headers = externalSheet.getRange("5:5").getValues()[0]; // Получаем строки заголовков

    for (var i = 0; i < students.length; i++) {
        var studentName = students[i][0].toString();
        if (studentName == "") continue;
        if (studentName == "total") break;
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
                    if ((headers[k+2] == "" || headers[k+2] == "PROJECT" || headers[k+2] == "Monday morning") && k < 5) continue;
                    if (headers[k+2] == "Training") {
                        trainingHours += data[j][k];
                        Logger.log("Training " + trainingHours);
                    }
                    else if (headers[k+2] == "HR") {
                        hrHours += data[j][k];
                        Logger.log("HR " + hrHours);
                    }
                    else if (headers[k+2] == "") {
                        continue;
                    }
                }
            }
        }
        // Записываем часы обучения и HR часы в соответствующие колонки
        if(trainingHours > 0) sheet.getRange("G" + (i + 2)).setValue(trainingHours);
        else sheet.getRange("G" + (i + 2)).setValue("");
        if(hrHours > 0) sheet.getRange("J" + (i + 2)).setValue(hrHours);
        else sheet.getRange("J" + (i + 2)).setValue("");

        // // Записываем собранные данные в колонку O
        // if (projectHours != "") {
        //   sheet.getRange("M" + (i + 2)).setValue(projectHours);
        // } else {
        //   sheet.getRange("M" + (i + 2)).setValue("");
        // }
    }
}


function updateStudents() {

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
    var learningValues = [];

    for (var i = 1; i < currentData.length; i++) {
        var currentName = currentData[i][1].toString(); // Значение из колонки B текущего листа
        if(currentName == "") continue;

        for (var j = 0; j < targetData.length; j++) {
            var targetName = targetData[j][0].toString(); // Значение из колонки A целевого листа
            var targetValue = targetData[j][3]; // Значение из колонки D целевого листа
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
                    learningValues.push(ticketValue + " (" + targetValue.toFixed(2) + ")");
                } else {
                    if (targetKey.startsWith("LEARN")) learningValues.push(targetKey + " (" + targetValue.toFixed(2) + ")");
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

        if(learningValues.length > 0) currentSheet.getRange("C" + (i + 1)).setValue(learningValues.join(', ')); // заполняем колонку C найденными значениями
        else currentSheet.getRange("C" + (i + 1)).setValue("");

        // if(sum>0) currentSheet.getRange("P" + (i + 1)).setValue(sum);
        // else currentSheet.getRange("P" + (i + 1)).setValue("");

        sumH = 0; // Сбрасываем сумму для следующей строки
        sumK = 0; // Сбрасываем сумму для следующей строки
        sum = 0; // Сбрасываем сумму для следующей строки
        foundValues = []; // очищаем массив найденных значений для следующей строки
        learningValues = [];
    }
}

function updateStudentDataFromSalesReport() {
    var sourceSpreadsheetId = '1CeYe0hb97nqDcOy9c-JRpW3wIvV6hr99IwElIyHv3fc'; // ID исходного документа
    var sourceSheetName = 'ALL report last week'; // Укажите точное название листа в исходном документе
    var targetSheetName = 'Students'; // Название листа в текущем документе
    var startRow = 7; // Строка, с которой начинается обработка в исходном документе
    var headerRow = 5; // Строка с заголовками колонок

    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
    var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);

    // Чтение заголовков колонок
    var headers = sourceSheet.getRange('B' + headerRow + ':CC' + headerRow).getValues()[0];
    var planIndex = headers.indexOf("Plan");
    var factIndex = headers.indexOf("Fact");
    var totalPlanIndex = headers.indexOf("TOTAL plan");
    var totalFactIndex = headers.indexOf("TOTAL fact");

    var sourceData = sourceSheet.getRange('B' + startRow + ':CC' + sourceSheet.getLastRow()).getValues();
    var targetData = targetSheet.getRange('A2:N' + targetSheet.getLastRow()).getValues();

    for (var i = 0; i < sourceData.length; i++) {
        var sourceName = sourceData[i][0]; // Имя разработчика из исходного документа
        var plan = sourceData[i][planIndex]; // Значение Plan
        var fact = sourceData[i][factIndex]; // Значение Fact
        var totalPlan = sourceData[i][totalPlanIndex]; // Значение Plan
        var totalFact = sourceData[i][totalFactIndex]; // Значение Fact

        for (var j = 0; j < targetData.length; j++) {
            var targetName = targetData[j][0]; // Имя студента из текущего документа
            if (targetName.startsWith(sourceName)) {
                targetSheet.getRange('M' + (j + 2)).setValue(plan); // Записываем Plan в колонку M
                targetSheet.getRange('N' + (j + 2)).setValue(fact); // Записываем Fact в колонку N
                targetSheet.getRange('O' + (j + 2)).setValue(totalPlan);
                targetSheet.getRange('P' + (j + 2)).setValue(totalFact);
            }
        }
    }
}

function updateWeekFact() {
    updateSheet("Week report All", "H"); // , "N"
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

function updateStudentStack() {
    var sourceSpreadsheetId = '1CeYe0hb97nqDcOy9c-JRpW3wIvV6hr99IwElIyHv3fc';
    var sourceSheetName = 'ALL report'; // Замените на имя листа в исходном документе
    var studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");

    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

    var studentsData = studentsSheet.getDataRange().getValues();
    var sourceData = sourceSheet.getRange('B5:Z' + sourceSheet.getLastRow()).getValues(); // Предполагаем, что данные начинаются с 7 строки


    var startColumnIndex = 22; // Колонка V
    var totalColumns = studentsSheet.getMaxColumns(); // Общее количество колонок в листе

    // Очистка колонки V и каждой четвертой колонки после нее
    for (var col = startColumnIndex; col <= totalColumns; col += 4) {
        var rangeToClear = studentsSheet.getRange(2, col, studentsData.length - 1);
        rangeToClear.clearContent();
        rangeToClear.setBackground("#00008b"); // Установка темно-синего фона
    }

    for (var i = 0; i < studentsData.length; i++) {
        var studentName = studentsData[i][0]; // Имя студента в колонке B

        for (var j = 0; j < sourceData.length; j++) {
            var developerName = sourceData[j][0]; // Имя разработчика в колонке B исходного листа
            console.log(studentName + ' - ' + developerName);
            if (studentName.toLowerCase() === developerName.toLowerCase()) {
                // Найдено совпадение имени
                for (var k = 9; k < sourceData[j].length; k++) { // Идем по колонкам стека
                    if (sourceData[0][k] === "") break; // Останавливаемся на первой пустой ячейке в 7-й строке

                    var stackName = sourceData[0][k] ? sourceData[0][k].toLowerCase() : null; // Имя стека в нижнем регистре
                    var stackCompetence = sourceData[j][k] ? sourceData[j][k] : null; // Компетенция разработчика в стеке

                    console.log(stackName + ' ' + stackCompetence);

                    if(!stackName) continue;

                    // Найти соответствующую колонку в листе Students
                    var columnIndex = getMatchingStackColumnIndex(studentsSheet, stackName, 3); // 3 - смещение на три колонки вправо
                    if (columnIndex > -1) {


                        // Установка цвета в зависимости от уровня компетенции
                        var cell = studentsSheet.getRange(i + 1, columnIndex + 1);
                        var stackLevel = sourceData[j][k];
                        if (stackLevel.startsWith('jun')) {
                            cell.setBackground("#add8e6");  // Цвет для Junior (Светло-синий)
                        } else if (stackLevel.startsWith('mid')) {
                            cell.setBackground("#90ee90");  // Цвет для Middle (Светло-зелёный)
                        } else if (stackLevel.startsWith('sr')) {
                            cell.setBackground("#f4a460");  // Цвет для Senior (Светло-коричневый)
                        }

                        cell.setValue(stackCompetence).setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
                    }
                }
                break; // Переходим к следующему студенту
            }
        }
    }
}

function getMatchingStackColumnIndex(sheet, targetName, offset) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
        if (headers[i].toLowerCase() === targetName.toLowerCase()) {
            return i + offset;
        }
    }
    return -1; // Если не найдено совпадение
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
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var studentsSheet = spreadsheet.getSheetByName("Students");
    var learningTotalsSheet = spreadsheet.getSheetByName("Learning totals");
    var ticketsSheet = spreadsheet.getSheetByName("Tickets");

    var studentsData = studentsSheet.getDataRange().getValues();
    var learningTotalsData = learningTotalsSheet.getDataRange().getValues();
    var ticketsData = ticketsSheet.getDataRange().getValues();

    // Создание словаря тикетов
    var ticketsDict = {};
    for (var i = 0; i < ticketsData.length; i++) {
        var ticketKey = ticketsData[i][1]; // Предполагаем, что код тикета в колонке A
        var ticketValue = ticketsData[i][0]; // Предполагаем, что название тикета в колонке B
        ticketsDict[ticketKey] = ticketValue;
    }

    var totalColumnIndex = studentsData[0].indexOf("Total") + 1; // +1 для индекса в Google Sheets

    // Очистка комментариев в колонке Totals
    studentsSheet.getRange(2, totalColumnIndex, studentsData.length - 1).clearNote();

    for (var i = 1; i < studentsData.length; i++) {
        var studentName = studentsData[i][1];
        var totalHours = 0;
        var ticketDetails = [];

        if (studentName !== "") {
            for (var j = 1; j < learningTotalsData.length; j++) {
                var learningTotalsName = learningTotalsData[j][0];
                var ticketCode = learningTotalsData[j][1]; // Код тикета
                var hours = learningTotalsData[j][2];

                if (studentName === learningTotalsName && hours !== "") {
                    totalHours += hours;
                    var ticketTitle = ticketsDict[ticketCode] || ticketCode; // Заголовок тикета или код, если заголовок не найден
                    ticketDetails.push(ticketTitle + " (" + hours.toFixed(2) + " ч)");
                }
            }

            // Запись общего количества часов и добавление комментария с деталями тикетов
            var cell = studentsSheet.getRange(i + 1, totalColumnIndex);
            cell.setValue(totalHours);
            if (ticketDetails.length > 0) {
                cell.setNote("Тикеты:\n" + ticketDetails.join("\n"));
            }
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


