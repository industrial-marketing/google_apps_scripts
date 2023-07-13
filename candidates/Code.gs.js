function getFormulasForRow(row) {
    return {
        "BM": '=getEventDetailsForCandidate(A' + row + ')',
        "AL": '=parseConclusion(ROW(AJ' + row + '); COLUMN(AJ' + row + '))',
        "AG": '=parseEventDate($BM' + row + ')',
        "AH": '=parseEventTime($BM' + row + ')',
        "AF": '=IF(AG' + row + ' = "";"";"SharpDev Interview "&TEXT(AG' + row + ';"dd/mm/yyyy")&" "&IF(AH' + row + ' = "";"15:00";AH' + row + ')&"MSK")',
        "AI": '=updateInterviewStatus($BM' + row + ')',
        "AK": '=parseEventAttendees($BM' + row + ')',
        "BC": '=parseCustom($BC$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        "BD": '=parseCustom($BD$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        "BE": '=parseCustom($BE$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        "BF": '=parseCustom($BF$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        "BG": '=parseCustom($BG$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        "BH": '=parseCustom($BH$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        "BI": '=parseCustom($BI$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        "BJ": '=parseCustom($BJ$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        "BK": '=parseCustom($BK$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        "BL": '=parseCustom($BL$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        "H": '=parseEventComments(BM' + row + ';AL' + row + ';BB' + row + ')'
    };
}

function saveOAuthToken() {
    var token = ScriptApp.getOAuthToken();
    PropertiesService.getScriptProperties().setProperty('OAUTH_TOKEN', token);

}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
        .addItem('Save OAuth Token', 'saveOAuthToken')
        .addItem('Parse summary from "I" and fill with data', 'parseDataCheckOAuth')
        .addItem('Convert formulas to text', 'removeFunctions')
        .addItem('Convert text to formulas', 'replaceContentWithFormulasCheckOAuth')
        .addItem('Generate report for Current Week', 'generateReportForCurrentWeek')
        .addItem('Generate report for Previous Week', 'generateReportForPreviousWeek')
        .addItem('Generate report for Current Month', 'generateReportForCurrentMonth')
        .addItem('Generate report for Previous Month', 'generateReportForPreviousMonth')
        .addToUi();
}

function parseDataCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию parseData()
        parseData();
    } else {
        // Токен отсутствует, отображаем диалоговое окно пользователю
        var response = Browser.msgBox(
            "OAuth Token Required",
            "Please obtain an OAuth token by clicking the 'OK' button.",
            Browser.Buttons.OK_CANCEL
        );

        if (response === Browser.Buttons.OK) {
            // Пользователь нажал OK, выполняем действия для получения токена
            saveOAuthToken()
        } else {
            // Пользователь нажал Cancel, не выполняем функцию parseData()
            return;
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

function replaceContentWithFormulasCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию parseData()
        replaceContentWithFormulas();
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

function checkFileAvailability(fileId = '1cHGKFLggYyYUTTQGD_jB8cItB8oQ6jdxiC7XNfKG8C0') {
    var token = PropertiesService.getScriptProperties().getProperty('OAUTH_TOKEN');
    var url = "https://docs.googleapis.com/v1/documents/" + fileId;
    var options = {
        method: "GET",
        headers: {
            Authorization: "Bearer " + token
        },
        muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    Logger.log(statusCode);
    return statusCode === 200;
}

function parseData() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var row = sheet.getActiveCell().getRow();
    var cellContent = sheet.getRange(row, 9).getValue(); // получение содержимого ячейки из колонки I текущей строки

    // Регулярные выражения для шаблонов
    var namePattern = /^(.+?)\s*\(/;
    var stackPattern = /\((.+)\s*\)/;
    var salaryPattern = /(ЗП|запрос)\s*(.+)\s/i;
    var experiencePattern = /опыт\s*(.+)\s/;
    var filePattern = /(файл|file|interview|интервью)\s*(https:\/\/(?:drive|docs).google.com\/(?:file\/d\/|document\/d\/)\S+)\s/i;
    var cvPattern = /(c\/v|с\/v|cv|резюме)\s*(https:\/\/(?:drive|docs).google.com\/(?:file\/d\/|document\/d\/)\S+)\s/i;
    var telegramPattern = /(telegram|телеграм)\s*(.+?)\s/i || /(@)\s*(.+?)\s/i;
    var skypePattern = /(skype|скайп)\s*(.+?)\s/i;

    var nameMatch = namePattern.test(cellContent);
    var stackMatch = stackPattern.test(cellContent);
    var salaryMatch = salaryPattern.test(cellContent);
    var experienceMatch = experiencePattern.test(cellContent);
    var fileMatch = filePattern.test(cellContent);
    var cvMatch = cvPattern.test(cellContent);
    var telegramMatch = telegramPattern.test(cellContent);
    var skypeMatch = skypePattern.test(cellContent);

    // Проверка, что содержимое ячейки соответствует шаблонам
    if (
        nameMatch && stackMatch && salaryMatch && experienceMatch && fileMatch && cvMatch && (telegramMatch || skypeMatch)
    ) {
        // Извлекаем данные из cellContent с помощью регулярных выражений
        var name = cellContent.match(namePattern);
        var stack = cellContent.match(stackPattern);
        var salary = cellContent.match(salaryPattern);
        var experience = cellContent.match(experiencePattern);
        var file = cellContent.match(filePattern);
        var cv = cellContent.match(cvPattern);
        var telegram = cellContent.match(telegramPattern);
        var skype = cellContent.match(skypePattern);

        // Verify the validity of the Google Drive links
        var fileIdFile = extractFileId(file[2]);
        var fileIdCv = extractFileId(cv[2]);


        if (file) {
            var fileAvailable = checkFileAvailability(fileIdFile);
            if (!fileAvailable) {
                Browser.msgBox("Please check the link to Interview. It seems to be incorrect or the file is not accessible.");
                return; // Завершаем функцию
            }
        }

        if (!fileIdCv) {
            Browser.msgBox("Please check the link to CV. It seems to be incorrect or the file is not accessible.");
            return; // Завершаем функцию
        }


        // Создаем массив output и заполняем его данными
        var output = [];
        let date = new Date();
        let today = date.toLocaleDateString("ru-RU", { day: "2-digit", month: "2-digit", year: "numeric" });

        output[0] = name ? name[1] : ""; // A - имя
        output[1] = "Дарья"; // B - имя рекрутера
        output[2] = "Ожидает интервью"; // C - статус
        output[3] = today; // D - дата
        output[6] = stack ? stack[1] : ""; // G - стек
        output[9] = experience ? experience[1] : ""; // J - опыт
        output[12] = salary ? salary[2] : ""; // M - зарплатные ожидания
        output[11] = cv ? "=HYPERLINK(\"" + cv[2] + "\"; \"Резюме\")" : ""; // L - резюме
        output[35] = file ? "=HYPERLINK(\"" + file[2] + "\"; \"Interview\")" : ""; // AJ - файл интервью
        output[30] = "telegram: " + (telegram ? telegram[2] : "") + "\n" + "skype: " + (skype ? skype[2] : ""); // AE - контакты

        docId = parseDocID(file[2]);
        var fields = parseInterview(docId);
        fields = JSON.parse(fields);
        output[13] = fields.educationText;
        output[10] = fields.englishText;
        output[9] = output[9] + "\n" + fields.technologyText;
        output[8] = fields.text;

        // Обновляем ячейки текущей строки данными из output
        for (let i = 0; i < output.length; i++) {
            if (output[i] !== undefined) {
                sheet.getRange(row, i+1).setValue(output[i]);
            }
        }
        replaceContentWithFormulas();
        // добавить установку формата даты в дату инт
    } else {
        // Если содержимое ячейки не соответствует требуемому формату, выводится сообщение об ошибке
        var errorMessages = [];
        if (!nameMatch) errorMessages.push("name");
        if (!stackMatch) errorMessages.push("stack");
        if (!salaryMatch) errorMessages.push("salary");
        if (!experienceMatch) errorMessages.push("experience");
        if (!fileMatch) errorMessages.push("file");
        if (!cvMatch) errorMessages.push("CV");
        if (!telegramMatch && !skypeMatch) errorMessages.push("telegram or skype");

        Browser.msgBox("There were errors with the following fields: " + errorMessages.join(", "));
    }
}


function replaceContentWithFormulas() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange(); // получает выбранный диапазон
    var row = range.getRow(); // получает номер первой строки выбранного диапазона

    // Create a dictionary matching column name to formula
    var formulas = getFormulasForRow(row);

    // Iterate over each entry in the formulas dictionary and set the formula for the cell
    for (var column in formulas) {
        var cell = sheet.getRange(column + row);
        cell.setFormula(formulas[column]);
    }
    // Конвертируем колонку AG в формат даты
    var agCell = sheet.getRange("AG" + row);
    agCell.setNumberFormat("dd/mm/yyyy");
}

// Функция для генерации отчета
function generateReport(period = "currentWeek") {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Получаем исходные данные из листа "Candidates"
    var candidatesSheet = spreadsheet.getSheetByName("Candidates");
    var candidatesData = candidatesSheet.getDataRange().getValues();

    // Получаем текущую дату
    var now = new Date();

    // Получаем дату начала периода в зависимости от выбранного периода
    var periodStartDate;
    switch (period) {
        case 'currentWeek':
            periodStartDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay() + 1);
            periodEndDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay() + 7);
            break;
        case 'previousWeek':
            periodStartDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay() - 6);
            periodEndDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay());
            break;
        case 'currentMonth':
            periodStartDate = new Date(now.getFullYear(), now.getMonth(), 1);
            periodEndDate = new Date(now.getFullYear(), now.getMonth() + 1, 0);
            break;
        case 'previousMonth':
            periodStartDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
            periodEndDate = new Date(now.getFullYear(), now.getMonth(), 0);
            break;
        default:
            // Если период не указан или некорректен, используем текущую неделю
            periodStartDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay() + 1);
            periodEndDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay() + 7);
    }

    var reportData = [];
    var statusData = {};

    // Фильтруем данные с выбранного периода и выбираем нужные колонки
    for (var i = 0; i < candidatesData.length; i++) {
        var row = candidatesData[i];
        //if (row[1] != "Natalia Gromova") continue;  // если мы хотим получить отчет только по выбранному рекрутеру.
        var date = new Date(row[32]); // AG column is 33rd (index 32 as 0-based)
        var status = row[2];

        if (date >= periodStartDate && date <= periodEndDate) {
            var dataRow = [
                date,   // Column B
                row[0], // Column C
                "(" + row[6] + ")", // Column D
                row[1], // Column E
                "[" + row[2] + "]", // Column F
                row[7]  // Column G
            ];
            reportData.push(dataRow);

            // Запоминаем статусы для подсчета статистики
            if (!statusData[status]) {
                statusData[status] = [];
            }
            statusData[status].push(row[0] + " (" + row[6] + ")"); // Добавляем имя и стек
        }
    }

    // Сортировка данных по дате
    reportData.sort(function(a, b) {
        return a[0] - b[0];
    });

    // Добавляем данные в лист "Reports"
    var reportsSheet = spreadsheet.getSheetByName("Reports");
    var lastRow = reportsSheet.getLastRow();

    var startRow = lastRow + 5; // Отступ в 5 строк перед линией

    // Создание жирной горизонтальной линии
    var lineRange = reportsSheet.getRange(startRow, 1, 1, reportsSheet.getLastColumn());
    lineRange.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    lineRange.setFontWeight("bold");

    startRow += 5;

    // Устанавливаем заголовок отчета
    var reportTitle = "Рекрутинг с " + Utilities.formatDate(periodStartDate, Session.getScriptTimeZone(), "dd.MM") +
        " до " + Utilities.formatDate(periodEndDate, Session.getScriptTimeZone(), "dd.MM");

    var cell = reportsSheet.getRange(startRow, 2);
    cell.setValue(reportTitle);

    var textStyle = {
        fontSize: 20,
        bold: true
    };
    cell.setFontSize(textStyle.fontSize);
    cell.setFontWeight(textStyle.bold);

    // Вставка данных отчета
    var currentRow = startRow + 1; // Текущая строка для вставки данных
    var previousDate = null; // Предыдущая дата
    var interviewCount = 0; // Счетчик проведенных собеседований
    for (var i = 0; i < reportData.length; i++) {
        var row = reportData[i];
        var currentDate = row[0];

        // Проверяем, если текущая дата отличается от предыдущей, то вставляем дату в столбец B
        if (!previousDate || currentDate.getTime() !== previousDate.getTime()) {
            reportsSheet.getRange(currentRow, 2).setValue(currentDate);
            currentRow++;
        }

        // Вставляем данные кандидата в соответствующие столбцы
        reportsSheet.getRange(currentRow, 3).setValue(row[1]);
        reportsSheet.getRange(currentRow, 4).setValue(row[2]);
        reportsSheet.getRange(currentRow, 5).setValue(row[3]);
        reportsSheet.getRange(currentRow, 6).setValue(row[4]);
        reportsSheet.getRange(currentRow, 7).setValue(row[5]);

        currentRow++;

        // Обновляем предыдущую дату
        previousDate = currentDate;

        // Увеличиваем счетчик проведенных собеседований, исключая статусы "Самоотвод" и "Ожидает интервью"
        if (row[4] !== "Отмена" && row[4] !== "Самоотвод" && row[4] !== "Ожидает интервью") {
            interviewCount++;
        }
    }

    // Выводим подзаголовок "Проведено N собеседований"
    var interviewsText = getInterviewsText(interviewCount);
    var lastRow = reportsSheet.getLastRow();
    reportsSheet.getRange(lastRow + 2, 2).setValue("Проведено " + interviewsText);

    // Выводим сгруппированные данные по статусам
    var statusRow = lastRow + 4; // Оставляем пустую строку после подзаголовка
    for (var status in statusData) {
        var candidates = statusData[status];
        reportsSheet.getRange(statusRow, 2).setValue(status + " (" + candidates.length + ")");
        reportsSheet.getRange(statusRow, 3).setValue(candidates.join(", "));
        statusRow++;
    }

    // Копирование данных из диапазона C-G в B-F в строках, где столбец B пустой
    var reportRange = reportsSheet.getRange(startRow + 1, 2, currentRow - startRow - 1, 6);
    var reportValues = reportRange.getValues();
    for (var i = 0; i < reportValues.length; i++) {
        var row = reportValues[i];
        if (row[0] === "") {
            reportsSheet.getRange(startRow + 1 + i, 2, 1, 5).setValues([row.slice(1)]);
        }
    }

    // Очистка колонки G в отчете
    reportsSheet.getRange(startRow + 1, 7, currentRow - startRow - 1, 1).clearContent();
}

function getInterviewsText(count) {
    if (
        count === 1 ||
        (count % 10 === 1 && count % 100 !== 11)
    ) {
        return count + " собеседование";
    } else if (
        (count % 10 === 2 || count % 10 === 3 || count % 10 === 4) &&
        !(count % 100 >= 11 && count % 100 <= 14)
    ) {
        return count + " собеседования";
    } else {
        return count + " собеседований";
    }
}

function generateReportForCurrentWeek() {
    generateReport('currentWeek');
}

function generateReportForPreviousWeek() {
    generateReport('previousWeek');
}

function generateReportForCurrentMonth() {
    generateReport('currentMonth');
}

function generateReportForPreviousMonth() {
    generateReport('previousMonth');
}

function clearRowAndAddFormula() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var range = sheet.getActiveRange(); // Get currently selected range
    var row = range.getRow();

    var cellContent = sheet.getRange('I' + (row+1)).getValue(); // Get the content of the cell

    // Regular expressions for the patterns
    var namePattern = /^(.+?)\s*\(/;
    var stackPattern = /\((.+)\s*\)/;
    var salaryPattern = /(ЗП|запрос)\s*(.+)\s/i;
    var experiencePattern = /опыт\s*(.+)\s/;
    var filePattern = /(файл|file|interview|интервью)\s*(.+?)\s/i;
    var cvPattern = /(c\/v|cv|резюме)\s*(.+?)\s/i;
    var telegramPattern = /(telegram|телеграм)\s*(.+?)\s/i || /(@)\s*(.+?)\s/i;
    var skypePattern = /(skype|скайп)\s*(.+?)\s/i;

    // Check if the content matches the patterns
    if (
        namePattern.test(cellContent) &&
        stackPattern.test(cellContent) &&
        salaryPattern.test(cellContent) &&
        experiencePattern.test(cellContent) &&
        filePattern.test(cellContent) &&
        cvPattern.test(cellContent) &&
        (telegramPattern.test(cellContent) || skypePattern.test(cellContent))
    ) {
        // Clear the content of the entire row
        sheet.getRange(row, 1, 1, sheet.getLastColumn()).clearContent();

        // Set the formula in the first cell
        sheet.getRange(row, 1).setFormula('=parseData($I' + (row+1) + ')');
    } else {
        SpreadsheetApp.getUi().alert('The content of the cell does not match the required format. Please check it.');
    }
}

function removeFunctions() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange(); // получает выбранный диапазон
    var row = range.getRow(); // получает номер первой строки выбранного диапазона

    // создает словарь с соответствием между именем столбца и формулой
    var formulas = getFormulasForRow(row);
    var problemColumns = [];

    // создает словарь с соответствием между именем столбца и формулой
    // перебирает все столбцы в словаре
    for (var col in formulas) {
        var cell = sheet.getRange(col + row); // получает ячейку по A1-нотации

        // Проверьте, выполнилась ли формула без ошибок
        if (cell.getValue() === '#ERROR!' || cell.getValue() === '#N/A') {
            problemColumns.push(col);
        }
    }

    // Если есть проблемные столбцы, показать сообщение об ошибке
    if (problemColumns.length > 0) {
        Browser.msgBox('Please check the formulas in columns: ' + problemColumns.join(", ") + '. They seem to be still working or returning errors.');
    } else {
        try {
            // Перебирает все столбцы в словаре для преобразования формул в значения
            for (var col in formulas) {
                var cell = sheet.getRange(col + row); // получает ячейку по A1-нотации

                // Преобразуйте формулу в значение
                cell.copyTo(cell, {contentsOnly:true});
            }

            // Конвертируем колонку AG в формат даты
            var agCell = sheet.getRange("AG" + row);
            agCell.setNumberFormat("dd/mm/yyyy");

            // Все формулы успешно преобразованы в значения
            Browser.msgBox('Formulas have been removed successfully.');
        } catch (e) {
            // Выводим текст исключения
            Browser.msgBox('Exception: ' + e.message);
        }
    }
}


function updateFormulas() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange(); // получает выбранный диапазон
    var row = range.getRow(); // получает номер первой строки выбранного диапазона

    // Получите ячейку в столбце A для проверки
    var checkCell = sheet.getRange('A' + row);

    // Проверьте, содержит ли ячейка текстовые данные
    if (checkCell.getValue().toString() !== checkCell.getFormula()) {
        Browser.msgBox('Please check the row. You can fix hyperlinks and update interview/event only for the rows with formula in A column');
        return;
    }

    copyAndPasteValues(row);
    updateFormulasInRow(row);

    // создает словарь с соответствием между именем столбца и формулой
    var formulas = getFormulasForRow(row);

    // перебирает все столбцы в словаре
    for (var col in formulas) {
        var cell = sheet.getRange(col + row); // получает ячейку по A1-нотации
        cell.setFormula(formulas[col]); // устанавливает формулу для этой ячейки
    }
}

function copyAndPasteValues(row) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Получить диапазон выделенной строки
    var range = sheet.getRange(row, 1, 1, sheet.getLastColumn());

    // Скопировать и вставить значения ячеек без изменений
    range.copyTo(range, { contentsOnly: true });
}

function updateFormulasInRow(row) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Обновите формулу для ячейки с ссылкой на резюме
    var cvCell = sheet.getRange(row, 12);
    cvCell.setFormula(cvCell.getValue());

    // Обновите формулу для ячейки с ссылкой на файл интервью
    var interviewCell = sheet.getRange(row, 36);
    interviewCell.setFormula(interviewCell.getValue());
}

function parseEventDate(eventComments) {
    var date = "";

    var lines = eventComments.split('\n');
    if (lines.length > 0 && lines[0].trim() !== "") {
        var dateTimeLine = lines[0];
        var dateTimeParts = dateTimeLine.split(' ');
        if (dateTimeParts.length > 1) {
            var dateParts = dateTimeParts[0].split('.');
            if (dateParts.length === 3) {
                date = `${dateParts[0]}/${dateParts[1]}/${dateParts[2]}`;
            }
        }
    }

    return date;
}



function parseEventTime(eventComments) {
    var time = "";

    var lines = eventComments.split('\n');
    if (lines.length > 0 && lines[0].trim() !== "") {
        var dateTimeLine = lines[0];
        var dateTimeParts = dateTimeLine.split(' ');
        if (dateTimeParts.length > 1) {
            time = dateTimeParts[1];
        }
    }

    return time;
}

function parseEventAttendees(eventComments) {
    var attendees = "";
    var attendeesFound = false; // Флаг, указывающий на наличие строки "Attendees:"

    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; // Регулярное выражение для проверки email-адреса

    var lines = eventComments.split('\n');
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        if (line.startsWith("Attendees:")) {
            attendeesFound = true;
        } else if (attendeesFound && line.startsWith("Description:")) {
            break; // Прекращаем парсинг после строки "Description:"
        } else if (attendeesFound && emailRegex.test(line)) {
            attendees += line + "\n"; // Добавляем участников в список
        }
    }

    return attendees.trim(); // Удаляем возможные пробелы в начале и конце строки
}

function parseEventTitle(eventComments) {
    var title = "";

    var lines = eventComments.split('\n');
    if (lines.length > 0 && lines[0].trim() !== "") {
        var dateTimeLine = lines[0];
        var dateTimeParts = dateTimeLine.split(' ');
        if (dateTimeParts.length > 2) {
            title = dateTimeParts.slice(2).join(' ');
        }
    }

    return title;
}


function updateInterviewStatus(eventComments) {
    var status = "";

    var eventDate = parseEventDate(eventComments);
    var eventTime = parseEventTime(eventComments);
    var eventTitle = parseEventTitle(eventComments);

    var currentDate = new Date();
    var eventDateTime = new Date(eventDate + " " + eventTime);

    if (eventDateTime > currentDate) {
        status = "Назначено";
    } else {
        if (eventTitle.startsWith("Canceled:")) {
            status = "Отменено";
        } else {
            status = "Пройдено";
        }
    }

    return status;
}


function parseEventComments(eventComments, candidateScores, comments) {
    var parsedData = {};

    var lines = eventComments ? eventComments.split('\n') : [];

    // Проверяем, есть ли дата и время в первой строке
    if (lines.length > 0 && lines[0].trim() !== "") {
        var dateTimeLine = lines[0];
        var dateTimeParts = dateTimeLine.split(' ');
        if (dateTimeParts.length > 1) {
            var date = dateTimeParts[0];
            var time = dateTimeParts[1];
            parsedData.dateTime = date + ' в ' + time;
        }
    } else {
        parsedData.dateTime = 'no event yet';
    }

    // Ищем строку с причиной отмены
    var cancellationReason = "";
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        if (line.startsWith("Cancellation Reason:")) {
            cancellationReason = line.substring("Cancellation Reason:".length).trim();
            break;
        }
    }

    // Если найдена причина отмены, добавляем ее в результат
    if (cancellationReason !== "") {
        parsedData.cancellationReason = cancellationReason;
    }

    // Ищем строки с оценками кандидата
    var scoreLines = candidateScores ? candidateScores.split('\n') : [];
    scoreLines = scoreLines.filter(function(line) {
        return line.trim().startsWith("Оценка сумма") || line.trim().startsWith("Оценка техлида");
    });

    // Если есть оценки, добавляем их в результат
    if (scoreLines.length > 0) {
        parsedData.scores = scoreLines.join('\n');
    }

    // Формируем результат
    var result = "";
    if (parsedData.dateTime) {
        result += parsedData.dateTime;
        if (comments) {
            result += " " + comments;
        }
        if (parsedData.cancellationReason) {
            result += " " + parsedData.cancellationReason;
        }
        if (parsedData.scores) {
            result += "\n" + parsedData.scores;
        }
    }

    return result;
}



function getEventDetailsForCandidate(candidateName) {
    var now = new Date();
    var oneYearAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000); // 365 days for a year - исправили на 30 дней, нет смысла теперь год тянуть
    var fifteenDaysLater = new Date(now.getTime() + 15 * 24 * 60 * 60 * 1000);

    var oneYearAgoISOString = oneYearAgo.toISOString();
    var fifteenDaysLaterISOString = fifteenDaysLater.toISOString();

    var calendarId = 'primary'; // Use your calendar ID here
    var url = "https://www.googleapis.com/calendar/v3/calendars/" + calendarId + "/events?timeMin=" + oneYearAgoISOString + "&timeMax=" + fifteenDaysLaterISOString;

    var token = PropertiesService.getScriptProperties().getProperty('OAUTH_TOKEN');

    var options = {
        headers: {
            Authorization: 'Bearer ' + token
        }
    };

    var latestEventDetails = "";
    var latestEventDate = new Date(0);  // Initialize to earliest possible date
    var pageToken;
    do {
        var response = UrlFetchApp.fetch(url + (pageToken ? "&pageToken=" + pageToken : ""), options);
        var data = JSON.parse(response.getContentText());

        var events = data.items;

        for (var i=0; i<events.length; i++) {
            var event = events[i];
            var title = event.summary || "";

            if (title.toLowerCase().includes(candidateName.toLowerCase())) {
                var eventDate = new Date(event.start.dateTime); // Event time

                // If this event is later than the latest event found so far, update the event details
                if (eventDate.getTime() > latestEventDate.getTime()) {
                    latestEventDate = eventDate;
                    var formattedDate = Utilities.formatDate(eventDate, "Europe/Moscow", "dd.MM.yyyy HH:mm"); // Formatting the date and time according to Moscow timezone

                    var attendeesList = event.attendees && event.attendees.length > 0 ? event.attendees.map(function(a) { return a.email; }).join("\n") : "No attendees provided";
                    var description = event.description || "No description provided";
                    var meetLink = event.hangoutLink || "No Google Meet link provided";

                    latestEventDetails = formattedDate + "\n" + title + "\nAttendees:\n" + attendeesList + "\nDescription:\n" + description + "\nGoogle Meet: " + meetLink;
                }
            }
        }
        pageToken = data.nextPageToken;
    } while (pageToken);

    return latestEventDetails;
}

function extractFileId(url) {
    var match = url.match(/[-\w]{25,}/);
    return match ? match[0] : null;
}

function parseConclusion(row, column) {
    var docId = docidfromurl(row, column);
    var token = PropertiesService.getScriptProperties().getProperty('OAUTH_TOKEN');
    var url = "https://docs.googleapis.com/v1/documents/" + docId; // URL для запроса
    var text = "";
    var options = {
        "method": "get",
        "headers": {
            "Authorization": "Bearer " + token // добавить токен в заголовок
        },
        'fields': 'documentStyle,updateTime'
    };
    var response = UrlFetchApp.fetch(url, options); // выполнить запрос
    var document = JSON.parse(response.getContentText()); // получить JSON-объект документа
    var lastTable = null; // переменная для хранения последней таблицы
    var structuralElements = document.body.content; // массив structural elements в документе
    var totalEstimate = 0;
    var totalEstimates = 0;
    var partEstimate = {};
    var partEstimates = {};
    var currentKey = "general";

    // перебрать все structural elements в документе
    for (var a = 0; a < structuralElements.length; a++) {
        var element = structuralElements[a]; // текущий элемент
        if (element.paragraph) { // если элемент - абзац
            var paragraph = element.paragraph; // объект абзаца
            var textRuns = paragraph.elements; // массив text runs в абзаце

            // перебрать все text runs в абзаце
            for (var l = 0; l < textRuns.length; l++) {
                var textRun = textRuns[l]; // текущий text run
                if (textRun.textRun) var textContent = textRun.textRun.content; // текстовое содержимое text run
                if (textContent != 'undefined') {
                    var matchKey = checkStringForMatch(textContent);
                    if(matchKey) currentKey = matchKey;
                }
            }
        }
        if (element.table) { // если элемент - таблица
            var lastTable = element.table; // запомнить его как последнюю таблицу
            if (lastTable) { // если нашли хотя бы одну таблицу
                var tableRows = lastTable.tableRows; // массив table rows в таблице

                // перебрать все table rows в таблице
                for (var i = 0; i < tableRows.length; i++) {
                    var row = tableRows[i]; // текущая строка
                    var tableCells = row.tableCells; // массив table cells в строке

                    // перебрать все table cells в строке
                    for (var j = 0; j < tableCells.length; j++) {
                        var cell = tableCells[j]; // текущая ячейка
                        var cellElements = cell.content; // массив structural elements в ячейке

                        // перебрать все structural elements в ячейке
                        for (var k = 0; k < cellElements.length; k++) {
                            var cellElement = cellElements[k]; // текущий элемент

                            if (cellElement.paragraph) { // если элемент - абзац
                                var paragraph = cellElement.paragraph; // объект абзаца
                                var textRuns = paragraph.elements; // массив text runs в абзаце

                                // перебрать все text runs в абзаце
                                for (var l = 0; l < textRuns.length; l++) {
                                    var textRun = textRuns[l]; // текущий text run
                                    if (textRun.textRun) var textContent = textRun.textRun.content; // текстовое содержимое text run
                                    if (textContent != 'undefined') {
                                        //Logger.log("Найден текст: " + textContent);
                                        var estimate = textContent.match(/Оценка:\s*([0-5]{1})/); // искать в тексте "Оценка:"
                                        if (estimate || estimate == 0) { // если нашли "Оценка:"
                                            Logger.log("Найдена оценка: " + estimate[1] + " - всего оценок: " + totalEstimates + "сумма: " + totalEstimate);
                                            var estimate = +estimate[1];
                                            var totalEstimate = totalEstimate + estimate; // прибавить оценку к общей сумме
                                            if(!partEstimate[currentKey]) partEstimate[currentKey] = 0;
                                            partEstimate[currentKey] = partEstimate[currentKey] + estimate;

                                            if(!partEstimates[currentKey]) partEstimates[currentKey] = 0;
                                            partEstimates[currentKey] = partEstimates[currentKey] + 1;

                                            Logger.log(currentKey + " - " + partEstimate[currentKey]);
                                            var totalEstimates = totalEstimates + 1; // увеличить количество оценок на 1
                                        }
                                    }
                                }
                            }

                        }
                    }
                }
            }

        }
    }

    if (lastTable) { // если нашли хотя бы одну таблицу
        Logger.log("Найдена последняя таблица в документе.");
        var tableRows = lastTable.tableRows; // массив table rows в таблице

        // перебрать все table rows в таблице
        for (var i = 0; i < tableRows.length; i++) {
            var row = tableRows[i]; // текущая строка
            var tableCells = row.tableCells; // массив table cells в строке

            // перебрать все table cells в строке
            for (var j = 0; j < tableCells.length; j++) {
                var cell = tableCells[j]; // текущая ячейка
                var cellElements = cell.content; // массив structural elements в ячейке

                // перебрать все structural elements в ячейке
                for (var k = 0; k < cellElements.length; k++) {
                    var cellElement = cellElements[k]; // текущий элемент

                    if (cellElement.paragraph) { // если элемент - абзац
                        var paragraph = cellElement.paragraph; // объект абзаца
                        var textRuns = paragraph.elements; // массив text runs в абзаце

                        // перебрать все text runs в абзаце
                        for (var l = 0; l < textRuns.length; l++) {
                            var textRun = textRuns[l]; // текущий text run
                            var textContent = textRun.textRun.content; // текстовое содержимое text run
                            if (textContent != 'undefined') var text = text + textContent;
                            Logger.log("Найден текст: " + textContent); // вывести текст в лог
                        }
                    }

                    if (cellElement.inlineObjectElement) { // если элемент - inline object
                        var inlineObjectElement = cellElement.inlineObjectElement; // объект inline object
                        var inlineObjectId = inlineObjectElement.inlineObjectId; // идентификатор inline object

                        Logger.log("Найден inline object с ID: " + inlineObjectId); // вывести ID в лог
                    }
                }
            }
        }
    } else { // если не нашли ни одной таблицы
        Logger.log("В документе нет таблиц.");
    }

    if (totalEstimate > 0) {
        var totalEstimate = (totalEstimate/totalEstimates).toFixed(2);
        var total = "\nОценка суммарная: " + totalEstimate + "(всего оценок: " + totalEstimates + ")";
    }
    if (Object.keys(partEstimate).length > 0) {
        for (var key in partEstimate) {
            if (partEstimate.hasOwnProperty(key)) {
                var skillEstimate = (partEstimate[key] / partEstimates[key]).toFixed(2);
                total += "\nОценка техлида " + key + " : " + skillEstimate + " (всего оценок: " + partEstimates[key] + ")";
            }
        }
    }

    return text + total;

}

function checkStringForMatch(str) {
    var skillsMap = {
        "general": "Тех. собеседование",
        "netremove": ".NET (убрать",
        "netnew": ".Net (new, WIP)",
        "nodejs": "NodeJS",
        "react": "React/RN",
        "ios": "iOS Native",
        "android": "Android Native",
        "flutter": "Flutter",
        "angular": "Angular",
        "python": "Python",
        "django": "Django",
        "phpnew": "PHP (new)",
        "phpold": "PHP (old)",
        "laravel": "Laravel",
        "testing": "Testing",
        "analysis": "Analysis",
        "java": "Java",
        "multithreading": "Multithreading",
        "databases": "Базы данных",
        "apparchitecture": "Архитектура приложений",
        "devopsaws": "Devops (AWS only)",
        "devops": "Devops"
    };

    for (var key in skillsMap) {
        if (str.startsWith(skillsMap[key])) {
            return key;
        }
    }
    return null;
}

// Функция linkurl принимает номер строки и столбца в google spreadsheet и возвращает HTML-ссылку с url из этой ячейки
function docidfromurl(row, column) {
    // Получаем активный лист
    var sheet = SpreadsheetApp.getActiveSheet();
    // Получаем значение ячейки по номеру строки и столбца
    var richText = sheet.getRange(row, column).getRichTextValue();
    // Получаем гиперссылку из объекта RichTextValue
    var url = richText.getLinkUrl();
    var docID = parseDocID(url);
    return docID;
}


// Функция parseDocID принимает урл документа Google Docs и возвращает docID из него
function parseDocID(url) {
    // Создаем регулярное выражение для поиска docID
    var regex = /\/d\/(.*?)\/edit/;
    // Применяем регулярное выражение к урлу и получаем массив совпадений
    var matches = url.match(regex);
    // Проверяем, что массив совпадений не пустой
    if (matches && matches.length > 1) {
        // Возвращаем первое совпадение после \"/d/\", которое является docID
        return matches[1];
    } else {
        // Если массив совпадений пустой, возвращаем пустую строку
        return "";
    }
}


function parseSummary(row, column) {
    var docId = docidfromurl(row, column);
    var fields = JSON.parse(parseInterview(docId));
    return fields.text;
}

function parseEducation(row, column) {
    var docId = docidfromurl(row, column);
    var fields = JSON.parse(parseInterview(docId));
    return fields.educationText;
}

function parseEnglish(row, column) {
    var docId = docidfromurl(row, column);
    var fields = JSON.parse(parseInterview(docId));
    return fields.englishText;
}

function parseTechnology(row, column) {
    var docId = docidfromurl(row, column);
    var fields = JSON.parse(parseInterview(docId));
    return fields.technologyText;
}

function parseCustom(customKeyword, row, column) {
    var docId = docidfromurl(row, column);
    var fields = JSON.parse(parseInterview(docId,customKeyword));
    return fields.customText;
}

function parseInterview(docId,customKeyword) {
    // input - docId документа, который нужно распарсить
    if(!docId) var docId = "1mupGZIZXRcoje8MI2ekLZED5Q1lAZMLM9YE3_mWgKaY"; // идентификатор документа

    var token = PropertiesService.getScriptProperties().getProperty('OAUTH_TOKEN');
    var url = "https://docs.googleapis.com/v1/documents/" + docId; // URL для запроса
    var text = "";
    var options = {
        "method": "get",
        "headers": {
            "Authorization": "Bearer " + token // добавить токен в заголовок
        },
        'fields': 'documentStyle,updateTime'
    };
    var response = UrlFetchApp.fetch(url, options); // выполнить запрос
    var document = JSON.parse(response.getContentText()); // получить JSON-объект документа

    var fields = {};

    var lastTable = null; // переменная для хранения последней таблицы
    var structuralElements = document.body.content; // массив structural elements в документе
    var text = '';
    var technologyText = "";
    var englishText = "";
    var educationText = "";
    var customText = "";
    var current = "";

    // перебрать все structural elements в документе
    for (var a = 0; a < structuralElements.length; a++) {
        var element = structuralElements[a]; // текущий элемент
        var current;
        if (element.paragraph) { // если элемент - абзац
            var paragraph = element.paragraph; // объект абзаца
            var textRuns = paragraph.elements; // массив text runs в абзаце

            // перебрать все text runs в абзаце
            for (var l = 0; l < textRuns.length; l++) {
                var textRun = textRuns[l]; // текущий text run
                if (textRun.textRun) var textContent = textRun.textRun.content; // текстовое содержимое text run
                if (textContent != 'undefined') {
                    Logger.log("Найден текст: " + textContent);
                    if(!customKeyword && !stop) {
                        var stop = textContent.match(/(Самая сложная задача|Culture fit)\s*/); // Как нашли в тексте Самая сложная задача или Culture fit - останавливаем поиск
                    }
                    if(customKeyword && !customStop) {
                        var customStop = textContent.match(new RegExp(customKeyword + '\\s*'));  // Как нашли в тексте custom поле - останавливаем поиск
                    }
                    var keywords = {
                        "english": "Английский\\s*",
                        "experience": "Коммерческий опыт\\s*",
                        "education": "Высшее образование\\s*",
                        "remote": "Опыт удаленной работы\\s*",
                        "computer": "Есть ли хороший домашний компьютер\\s*",
                        "quiz": "Quiz\\s*",
                        "softskills": "Софт-скиллы\\s*",
                        "tracker": "ТРЕКЕР и как обычно отслеживает время\\s*",
                        "technology": "Технологии\\s*",
                        "location": "ЛОКАЦИЯ\\/РЕЛОКАЦИЯ\\s*",
                        "legal": "ИП\\/САМОЗАНЯТЫЙ\\s*",
                        "business": "Есть ли у вас сейчас свой бизнес\\s*",
                        "daysoff": "Планируете ли вы отпуск\\s*",
                        "latestproject": "Последний проект\\s*",
                    };

                    if(customKeyword) {
                        keywords["custom"] = customKeyword + '\\s*';
                    }

                    var matches = {};
                    for(var key in keywords) {
                        matches[key] = textContent.match(new RegExp(keywords[key]));
                    }

                    for(var key in matches) {
                        if(matches[key]) {
                            text += matches[key][0];
                            Logger.log(matches[key][0]);
                            current = key;
                        }
                    }

                }
            }
        }
        var lastTable = element.table; // запомнить его как последнюю таблицу

        if (lastTable) { // если нашли хотя бы одну таблицу
            var tableRows = lastTable.tableRows;
            for (var i = 0; i < tableRows.length; i++) {
                var row = tableRows[i];
                var tableCells = row.tableCells;
                for (var j = 0; j < tableCells.length; j++) {
                    var cell = tableCells[j];
                    var cellElements = cell.content;
                    for (var k = 0; k < cellElements.length; k++) {
                        var cellElement = cellElements[k];
                        if (cellElement.paragraph) { // если элемент - абзац
                            var paragraph = cellElement.paragraph; // объект абзаца
                            var textRuns = paragraph.elements; // массив text runs в абзаце
                            for (var l = 0; l < textRuns.length; l++) {
                                var textRun = textRuns[l]; // текущий text run
                                if (textRun.textRun) var textContent = textRun.textRun.content; // текстовое содержимое text run
                                if (textContent != 'undefined') {
                                    var serviceField = textContent.match(/(Оценка|Ответ|ЗП)\s*/);
                                    if (!serviceField) {
                                        textContent = textContent.replace(/: \n/g, " ");
                                        Logger.log(textContent);
                                        if (current && current != "technology") text += textContent;

                                        switch (current) {
                                            case "technology":
                                                technologyText += textContent;
                                                break;
                                            case "education":
                                                educationText += textContent;
                                                break;
                                            case "english":
                                                englishText += textContent;
                                                break;
                                            case "custom":
                                                customText += textContent;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if(stop) {
            break;
        }
        if(customStop) {
            var stop = true;
        }
    }
    fields.text = replaceLinebreaks(text);
    fields.educationText = replaceLinebreaks(educationText);
    fields.englishText = replaceLinebreaks(englishText);
    fields.technologyText = replaceLinebreaks(technologyText);
    fields.customText = replaceLinebreaks(customText);
    return JSON.stringify(fields);
}

function replaceLinebreaks(text) {
    if (text) { // проверка на undefined
        return text.replace(/\n\n/g, '').replace(/:/g, '');
    }
    return ''; // возвращает пустую строку, если text не определен
}

function testGetInterviewsText() {
    for (var count = 1; count <= 1000; count++) {
        var interviewsText = getInterviewsText(count);
        Logger.log(count + ": " + interviewsText);
    }
}