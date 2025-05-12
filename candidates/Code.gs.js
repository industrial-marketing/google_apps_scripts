function getFormulasForRow(row) {
    fillFormulas(row);
    return {
        // "BM": '=getEventDetailsForCandidate(A' + row + ')',
        "E": '=AG' + row,
        "H": '=parseEventComments(BM' + row + ';AL' + row + ';BB' + row + ')',
        //"I": '=parseSummary(ROW(AJ' + row + '); COLUMN(AJ' + row + '))',
        //"J": '=parseTechnology(ROW(AJ' + row + '); COLUMN(AJ' + row + '))',
        //"K": '=parseEnglish(ROW(AJ' + row + '); COLUMN(AJ' + row + '))',
        //"N": '=parseEducation(ROW(AJ' + row + '); COLUMN(AJ' + row + '))',
        "AF": '=IF(AG' + row + ' = "";"";"SharpDev Interview "&TEXT(AG' + row + ';"dd/mm/yyyy")&" "&IF(AH' + row + ' = "";"15:00";AH' + row + ')&"MSK")',
        "AG": '=parseEventDate($BM' + row + ')',
        "AH": '=parseEventTime($BM' + row + ')',
        "AI": '=updateInterviewStatus($BM' + row + ')',
        "AK": '=parseEventAttendees($BM' + row + ')',
        "AL": '=BP' + row, //'=parseConclusion(ROW(AJ' + row + '); COLUMN(AJ' + row + '))',
        //"BC": '=parseCustom($BC$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        //"BD": '=parseCustom($BD$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        //"BE": '=parseCustom($BE$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        //"BF": '=parseCustom($BF$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        //"BG": '=parseCustom($BG$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        //"BH": '=parseCustom($BH$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        //"BI": '=parseCustom($BI$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        //"BJ": '=parseCustom($BJ$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        //"BK": '=parseCustom($BK$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        //"BL": '=parseCustom($BL$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',

        "BS": '=parseCustom($BS$1; ROW($AJ' + row + '); COLUMN($AJ' + row + '))',
        "BM": '=BO' + row,
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
        .addItem('Add candididates from Candidates Source/Screen', 'addCandidatesCheckOAuth')
        // .addItem('Update&Save row', 'updateAndSaveCheckOAuth')
        .addItem('Update row (add functions)', 'replaceContentWithFormulasCheckOAuth')
        .addItem('Save row (remove functions)', 'removeFunctions')
        .addItem('Get events from Google Calendar', 'updateCandidatesWithEvents')
        .addItem('Get conclusions from latest interview files', 'updateInterviewConclusions')
        .addItem('Generate Last/Current/Next reports', 'automatedReportGeneration')
        .addItem('Generate report for Current Week', 'generateReportForCurrentWeek')
        .addItem('Generate report for Last Week', 'generateReportForPreviousWeek')
        .addItem('Generate report for Current Month', 'generateReportForCurrentMonth')
        .addItem('Generate report for Last Month', 'generateReportForPreviousMonth')
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
            saveOAuthToken();
        } else {
            // Пользователь нажал Cancel, не выполняем функцию parseData()
            return;
        }
    }
}

function addCandidatesCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию parseData()
        addCandidatesFromSource();
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

function updateAndSaveCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию parseData()
        replaceContentWithFormulas();
        removeFunctions();
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

        sheet.setRowHeight(row, 400);


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


function getStandardDriveLink(url) {
    if (!url || !url.includes('id=') || !url.includes('&')) return false;
    var fileId = url.split('id=')[1].split('&')[0]; // Извлекаем ID файла
    var standardLink = "https://drive.google.com/file/d/" + fileId + "/view";
    return standardLink;
}


function addCandidatesFromSource() {
    var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Кандидаты сорс/скрин");
    var workflowSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Candidates");
    var sourceData = sourceSheet.getDataRange().getValues();
    var workflowData = workflowSheet.getDataRange().getValues();


    // Отметка времени	Адрес электронной почты	Имя кандидата на русском языке	Имя на английском языке	Резюме	Файл интервью	вакансия	стек	Статус	Комментарий статуса	Дата скрининга	Рекрутер скрининга	телеграм	Дата обновления статуса	город/страна	уровень английского	мин зп	max зп	Валюта	Skype	Email	Телефон	Комментарии	Результаты собеседования	Время скрининга	Сумма от	Сумма комфорт	Дубликаты	ID	Candidates status	Рекомендация от кого																																				

    // Названия колонок в листе исходных данных
    var headers = sourceData[0];
    var nameIndex = headers.indexOf("Имя кандидата на русском языке");
    var stackIndex = headers.indexOf("вакансия");
    var salaryMinIndex = headers.indexOf("мин зп");
    var salaryIndex = headers.indexOf("max зп"); //Сумма запроса (комфорт)
    var experienceIndex = headers.indexOf("Комментарий статуса");
    var commentsIndex = headers.indexOf("Комментарии");
    var fileIndex = headers.indexOf("Файл интервью");
    var cvIndex = headers.indexOf("Резюме");
    var telegramIndex = headers.indexOf("телеграм");
    var skypeIndex = headers.indexOf("Skype");
    var idIndex = headers.indexOf("ID");
    var recruiterIndex = headers.indexOf("Рекрутер скрининга");
    var locationIndex = headers.indexOf("город/страна");

    //console.log('Разбор данных со строки ' + (sourceData.length-30));

    for (var i = 1; i < sourceData.length; i++) {
        var row = sourceData[i];
        var status = row[headers.indexOf("Статус")];
        var candidatesStatus = row[headers.indexOf("Candidates status")];

        let cv = extractFileId(row[cvIndex]);
        if (!cv) {
            cv = getStandardDriveLink(row[cvIndex]);
            if (cv) row[cvIndex] = cv;
        }
        var file = extractFileId(row[fileIndex]);


        // var isCandidate = isInterviewAlreadyInWorkflow(workflowData, row[idIndex]);
        // var fileAvailable = checkFileAvailability(file);

        //console.log(row[nameIndex] + ' ' + isCandidate + ' ' + fileAvailable);

        console.log(status,row[nameIndex], row[stackIndex], row[salaryIndex],row[fileIndex],row[cvIndex],row[telegramIndex],row[skypeIndex],file, cv,candidatesStatus);
        if (
            status === "Ожидает интервью" &&
            row[nameIndex] && row[stackIndex] && row[salaryIndex] &&
            row[fileIndex] && row[cvIndex] && (row[telegramIndex] || row[skypeIndex]) &&
            file &&
            cv && candidatesStatus === ''
        ) {
            console.log(row);

            var isCandidate = isInterviewAlreadyInWorkflow(workflowData, row[idIndex]);
            var fileAvailable = checkFileAvailability(file);

            console.log(row[nameIndex] + ' ' + isCandidate + ' ' + fileAvailable);


            if (!isCandidate && fileAvailable) {

                // Добавление данных кандидата в "Candidates workflow"
                addToWorkflow(workflowSheet, row, nameIndex, stackIndex, salaryIndex, salaryMinIndex, experienceIndex, fileIndex, cvIndex, telegramIndex, skypeIndex, commentsIndex, idIndex, recruiterIndex, locationIndex);
            }
        }

    }
}

function addToWorkflow(sheet, rowData, nameIdx, stackIdx, salaryIdx, salaryMinIdx, expIdx, fileIdx, cvIdx, telIdx, skypeIdx, commentsIdx, idIdx, recruiterIdx, locationIdx) {
    var newRow = [];
    // let date = new Date();
    // let today = date.toLocaleDateString("ru-RU", { day: "2-digit", month: "2-digit", year: "numeric" });

    // Заполнение данных для новой строки
    newRow[0] = rowData[nameIdx]; // A - Имя кандидата
    newRow[1] = rowData[recruiterIdx]; // B - Имя рекрутера
    newRow[2] = "Ожидает интервью"; // C - Статус
    //newRow[3] = today; // D - Дата
    newRow[6] = rowData[stackIdx]; // G - Вакансия
    newRow[9] = rowData[expIdx] + ' ' + rowData[commentsIdx]; // J - Комментарий статуса
    newRow[12] = (rowData[salaryMinIdx] ? rowData[salaryMinIdx] + '-' : '') + rowData[salaryIdx]; // M - Сумма запроса
    newRow[11] = "=HYPERLINK(\"" + rowData[cvIdx] + "\"; \"Резюме\")"; // L - Резюме
    newRow[35] = "=HYPERLINK(\"" + rowData[fileIdx] + "\"; \"Interview\")"; // AJ - Файл интервью
    newRow[30] = "telegram: " + (rowData[telIdx] ? rowData[telIdx] : "") + "\n" + "skype: " + (rowData[skypeIdx] ? rowData[skypeIdx] : ""); // AE - Контакты
    newRow[69] = rowData[idIdx];
    newRow[71] = rowData[locationIdx];
    newRow[72] = '=IFERROR(HYPERLINK("https://docs.google.com/spreadsheets/d/189YZ_AKtBhVBADGksYIjKQCg8h_ky6Bh5tjEzxUWeXY/edit#gid=1941670349&range=A" & MATCH("' + rowData[idIdx] + '"; \'Кандидаты сорс/скрин\'!AC:AC; 0) & ":Z" & MATCH("' + rowData[idIdx] + '"; \'Кандидаты сорс/скрин\'!AC:AC; 0); INDEX(\'Кандидаты сорс/скрин\'!C:C; MATCH("' + rowData[idIdx] + '"; \'Кандидаты сорс/скрин\'!AC:AC; 0)))';

    var docId = parseDocID(rowData[fileIdx]); // Предполагаем, что parseDocID - это ваша функция
    var fields = parseInterview(docId); // Предполагаем, что parseInterview - это ваша функция
    fields = JSON.parse(fields);

    // Добавление дополнительных данных
    newRow[13] = fields.educationText;
    newRow[10] = fields.englishText;
    newRow[9] = newRow[9] + "\n" + fields.technologyText;
    newRow[8] = fields.text;

    // // Обновляем ячейки текущей строки данными из output
    // for (let i = 0; i < output.length; i++) {
    //   if (output[i] !== undefined) {
    //     sheet.getRange(row, i+1).setValue(output[i]);
    //   }
    // }
    //
    // Добавление новой строки в таблицу
    sheet.appendRow(newRow);
    var lastRow = sheet.getLastRow();
    replaceContentWithFormulas(lastRow);
    // removeFunctions(lastRow);

}

function validateRow(row, cv, file, isCandidate) {
    var errors = [];
    // Проверки для каждого поля, аналогично вашему примеру
    // Например, если имя кандидата пустое, добавляем ошибку
    var nameIndex = headers.indexOf("Имя кандидата на русском языке");
    var stackIndex = headers.indexOf("вакансия");
    var salaryIndex = headers.indexOf("мин зп");
    var experienceIndex = headers.indexOf("Комментарий статуса");
    var telegramIndex = headers.indexOf("телеграм");
    var skypeIndex = headers.indexOf("Skype");

    if (!row[nameIndex]) errorMessages.push("name");
    if (!row[stackIndex]) errorMessages.push("stack");
    if (!row[salaryIndex]) errorMessages.push("salary");
    if (!row[experienceIndex]) errorMessages.push("experience");
    if (!file) errorMessages.push("file");
    if (!cv) errorMessages.push("CV");
    if (!row[skypeIndex] && !row[telegramIndex]) errorMessages.push("telegram or skype");
    if (isCandidate) errors.push("Candidate was added");

    // Добавьте здесь другие проверки

    return errors;
}

function isInterviewAlreadyInWorkflow(workflowData, id) {

    for (var i = 1; i < workflowData.length; i++) {
        var cellContent = workflowData[i][69];
        if (cellContent !== '' && cellContent.includes(id)) {
            return true;
        }
    }
    return false;
}

function replaceContentWithFormulas(row) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange(); // получает выбранный диапазон
    if (!row) var row = range.getRow(); // получает номер первой строки выбранного диапазона

    // Получаем текст из ячейки H
    var cellHValue = sheet.getRange('H' + row).getValue().toString();

    // Получаем текст из ячейки BB
    var cellBBValue = sheet.getRange('BB' + row).getValue().toString();

    // Ищем индекс начала блока оценок
    var index = cellHValue.indexOf("Оценка суммарная:");

    // Обрезаем текст, если нашли индекс
    var textBeforeRatings = index !== -1 ? cellHValue.substring(0, index) : cellHValue;

    if (cellBBValue !== textBeforeRatings && cellBBValue !== '') {
        // Показываем предупреждение пользователю
        var ui = SpreadsheetApp.getUi(); // Same variations.
        var response = ui.alert('Внимание', 'При обновлении текст \n\n\n"' + cellBBValue + '"\n\n\n в ячейке BB будет заменен на текст \n\n\n"' + textBeforeRatings + '"\n\n\n из ячейки H. \n Продолжить?', ui.ButtonSet.YES_NO);

        // Проверяем ответ пользователя
        if (response == ui.Button.YES) {
            // Если пользователь согласен, обновляем значение в ячейке BB
            sheet.getRange('BB' + row).setValue(textBeforeRatings);
        } else {
            // Если пользователь отказался, прерываем выполнение функции
            return;
        }
    } else {
        // Если пользователь согласен, обновляем значение в ячейке BB
        sheet.getRange('BB' + row).setValue(textBeforeRatings);
    }

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

// Функция для генерации отчета
function generateReport(period = "currentWeek", outputSheetName = "Reports") {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Получаем исходные данные из листа "Candidates"
    var candidatesSheet = spreadsheet.getSheetByName("Candidates");
    var candidatesData = candidatesSheet.getDataRange().getValues();

    // Определяем даты начала и конца периода
    var now = new Date();
    var periodStartDate, periodEndDate;
    // Определение начала и конца периода (как в вашем исходном коде)

    // Фильтрация и подготовка данных для отчета (как в вашем исходном коде)

    // Сортировка и подготовка к выводу (как в вашем исходном коде)

    var baseUrl = spreadsheet.getUrl().split('/edit')[0]; // Базовый URL для создания ссылок



    switch (period) {
        case 'nextWeek':
            periodStartDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay() + 8);
            periodEndDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - now.getDay() + 14);
            break;
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
            var link = baseUrl + '/edit#gid=0&range=' + (i + 1) + ':' + (i + 1); // Создание ссылки на строку
            var dataRow = [
                date,
                '=HYPERLINK("' + link + '";"' + row[0] + ' (' + row[71] + ')' + '")', // Ссылка вместо просто имени
                "(" + row[6] + ")",
                row[1],
                "[" + row[2] + "]",
                row[7]
            ];
            reportData.push(dataRow);

            // Запоминаем статусы для подсчета статистики
            if (!statusData[status]) {
                statusData[status] = [];
            }
            statusData[status].push(row[0] + " (" + row[6] + ")"); // Добавляем имя и стек
        }
    }

    // // Сортировка данных по дате
    // reportData.sort(function(a, b) {
    //     return a[0] - b[0];
    // });

    // Сортировка данных по дате и по row[7], если даты совпадают
    reportData.sort(function(a, b) {
        var dateComparison = a[0] - b[0];
        if (dateComparison === 0) {
            return a[5].localeCompare(b[5]); // Сравнение строк
        }
        return dateComparison;
    });

    // Добавляем данные в лист для отчетов
    var reportsSheet = spreadsheet.getSheetByName(outputSheetName);
    if (!reportsSheet) {
        reportsSheet = spreadsheet.insertSheet(outputSheetName);
    } else if (outputSheetName !== 'Reports') {
        reportsSheet.clear(); // Очищаем лист перед добавлением новых данных
    }

    var lastRow = reportsSheet.getLastRow();

    var startRow = lastRow + 5; // Отступ в 5 строк перед линией

    if (outputSheetName === 'Reports') {
        // Создание жирной горизонтальной линии
        var lineRange = reportsSheet.getRange(startRow, 1, 1, reportsSheet.getLastColumn());
        lineRange.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        lineRange.setFontWeight("bold");

        startRow += 5;
    }

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
    // Стилизация дат
    var textStyleDate = { fontSize: 14, bold: true };

    // Вставка данных отчета
    var currentRow = startRow + 1; // Текущая строка для вставки данных
    var previousDate = null; // Предыдущая дата
    var interviewCount = 0; // Счетчик проведенных собеседований
    for (var i = 0; i < reportData.length; i++) {
        var row = reportData[i];
        var currentDate = row[0];

        // Проверяем, если текущая дата отличается от предыдущей, то вставляем дату в столбец B
        if (!previousDate || currentDate.getTime() !== previousDate.getTime()) {
            var dateCell = reportsSheet.getRange(currentRow, 2);
            dateCell.setValue(currentDate);
            dateCell.setFontSize(textStyleDate.fontSize)
            dateCell.setFontWeight(textStyleDate.bold)
            currentRow++;
        }

        // Вставляем данные кандидата в соответствующие столбцы
        reportsSheet.getRange(currentRow, 2).setValue(row[1]);
        reportsSheet.getRange(currentRow, 3).setValue(row[2]);
        reportsSheet.getRange(currentRow, 4).setValue(row[3]);
        reportsSheet.getRange(currentRow, 5).setValue(row[4]);
        reportsSheet.getRange(currentRow, 6).setValue(row[5]);

        currentRow++;
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

    if (reportData.length === 0) return;

    // Выводим сгруппированные данные по статусам
    var statusRow = lastRow + 4; // Оставляем пустую строку после подзаголовка
    for (var status in statusData) {
        var candidates = statusData[status];
        reportsSheet.getRange(statusRow, 2).setValue(status + " (" + candidates.length + ")");
        reportsSheet.getRange(statusRow, 3).setValue(candidates.join(", "));
        statusRow++;
    }
}

function automatedReportGeneration() {
    // Указываем, что отчет за прошлую неделю должен быть сгенерирован на листе LastWeekReport
    generateReportForSpecificSheet('previousWeek', 'LastWeekReport');
    // Указываем, что отчет за текущую неделю должен быть сгенерирован на листе CurrentWeekReport
    generateReportForSpecificSheet('currentWeek', 'CurrentWeekReport');
    // Указываем, что отчет за текущую неделю должен быть сгенерирован на листе NextWeekReport
    generateReportForSpecificSheet('nextWeek', 'NextWeekReport');
}

function generateReportForSpecificSheet(period, sheetName) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var targetSheet = spreadsheet.getSheetByName(sheetName);

    // Проверяем, существует ли лист, если нет, то создаем его
    if (!targetSheet) {
        spreadsheet.insertSheet(sheetName);
        targetSheet = spreadsheet.getSheetByName(sheetName);
    }

    // Очищаем содержимое листа перед генерацией нового отчета
    targetSheet.clear();

    // Переопределяем функцию для выбора исходных и целевых листов внутри generateReport
    var originalGetSheetByName = spreadsheet.getSheetByName;
    spreadsheet.getSheetByName = function(name) {
        if (name === 'Reports') {
            return targetSheet;
        } else {
            return originalGetSheetByName.apply(spreadsheet, arguments);
        }
    };

    // Генерация отчета
    generateReport(period, sheetName);

    // Восстанавливаем оригинальную функцию getSheetByName
    spreadsheet.getSheetByName = originalGetSheetByName;
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

function removeFunctions(row) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange(); // получает выбранный диапазон
    if (!row) var row = range.getRow(); // получает номер первой строки выбранного диапазона

    // создает словарь с соответствием между именем столбца и формулой
    var formulas = getFormulasForRow(row);
    var problemColumns = [];

    // создает словарь с соответствием между именем столбца и формулой
    // перебирает все столбцы в словаре
    for (var col in formulas) {
        var cell = sheet.getRange(col + row); // получает ячейку по A1-нотации

        // Проверьте, выполнилась ли формула без ошибок
        if (cell.getValue() === '#ERROR!' || cell.getValue() === '#N/A') {
            problemColumns.push(col + row + ': "' + cell.getValue() + '"\n');
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
            // Browser.msgBox('Formulas have been removed successfully.');
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
    var latestDate = null;
    var latestDateTimeLine = "";

    // Разделение текста на блоки событий и удаление пустых строк
    var events = eventComments.split('\n\n===============\n\n').filter(event => event.trim() !== "");

    events.forEach(event => {
        var lines = event.split('\n');
        if (lines.length > 0 && lines[0].trim() !== "") {
            var dateTimeLine = lines[0];
            var dateTimeParts = dateTimeLine.split(' ');
            if (dateTimeParts.length > 1) {
                var dateParts = dateTimeParts[0].split('.');
                if (dateParts.length === 3) {
                    var eventDate = new Date(`${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`);
                    if (!latestDate || eventDate > latestDate) {
                        latestDate = eventDate;
                        latestDateTimeLine = dateTimeLine;
                    }
                }
            }
        }
    });

    if (latestDate) {
        var dateParts = latestDateTimeLine.split(' ')[0].split('.');
        return `${dateParts[0]}/${dateParts[1]}/${dateParts[2]}`; // Возвращаем дату в формате дд/мм/гггг
    } else {
        return "";
    }
}

function parseEventTime(eventComments) {
    var latestDate = null;
    var latestTime = "";

    // Разделение текста на блоки событий и удаление пустых строк
    var events = eventComments.split('\n\n===============\n\n').filter(event => event.trim() !== "");

    events.forEach(event => {
        var lines = event.split('\n');
        if (lines.length > 0 && lines[0].trim() !== "") {
            var dateTimeLine = lines[0];
            var dateTimeParts = dateTimeLine.split(' ');
            if (dateTimeParts.length > 1) {
                var dateParts = dateTimeParts[0].split('.');
                if (dateParts.length === 3) {
                    var eventDate = new Date(`${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`);
                    if (!latestDate || eventDate > latestDate) {
                        latestDate = eventDate;
                        latestTime = dateTimeParts[1]; // Запоминаем время самого позднего события
                    }
                }
            }
        }
    });

    return latestTime;
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
            //if (emailRegex.test(line)) {
            attendees += line + "\n"; // Добавляем участников в список - новая версия - в одну строку с заголовком Attendees:
            //}
        } else if (attendeesFound && (line.startsWith("Google Meet:") || line.startsWith("Description:"))) {
            break; // Прекращаем парсинг после строки "Description:"
        } else if (attendeesFound && emailRegex.test(line)) {
            attendees += line + "\n"; // Добавляем участников в список - старая версия - с заголовком Attendees: сверху.
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
    if (eventComments.endsWith('\n\n===============\n\n')) {
        eventComments = eventComments.substring(0, eventComments.length - '\n\n===============\n\n'.length);
    }
    var eventsData = eventComments.split('\n\n===============\n\n');
    var parsedEvents = [];
    var commentsLines = comments ? comments.split('\n') : [];

    // Обрабатываем каждое событие отдельно
    eventsData.forEach(function(event) {
        var parsedData = {};
        var lines = event.split('\n');

        // Парсинг даты и времени события
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

        // Ищем строку с причиной отмены и добавляем ее, если нашли
        var cancellationReasonIndex = lines.findIndex(line => line.trim().startsWith("Cancellation Reason:"));
        if (cancellationReasonIndex !== -1) {
            parsedData.cancellationReason = lines[cancellationReasonIndex].substring("Cancellation Reason:".length).trim();
        }

        // Добавляем событие в массив
        parsedEvents.push(parsedData);
    });

    // Ассоциирование комментариев с событиями
    parsedEvents.forEach(function(event, index) {
        var eventDate = event.dateTime.split(' в ')[0];
        var foundComment = false;
        commentsLines.forEach(function(commentLine) {
            var commentDateTime = commentLine.split(' в ')[0]?.trim() + ' в ' + commentLine.split(' ')[2]?.trim();
            if (commentDateTime === event.dateTime.trim()) {
                foundComment = true;
                event.comment = commentLine.substring(commentDateTime.length + 1).trim();
            }
        });

        // Если комментарий не найден и существует только одно событие, используем весь текст комментария
        if (!foundComment && eventsData.length === 1 && index === 0) {
            event.comment = comments;
        }
    });

    // Удаление дублирующихся комментариев об отмене из комментариев BB, если они уже указаны в деталях события
    parsedEvents.forEach(function(event) {
        if (event.cancellationReason && event.comment && event.comment.includes(event.cancellationReason)) {
            event.comment = event.comment.replace(event.cancellationReason, "").trim();
        }
    });

    // Парсинг оценок (добавляем оценки только один раз, в конец результата)
    var scoreLines = candidateScores ? candidateScores.split('\n') : [];
    scoreLines = scoreLines.filter(function(line) {
        return line.trim().startsWith("Оценка суммарная") || line.trim().startsWith("Оценка техлида");
    });
    var scores = scoreLines.join('\n');

    // Функция преобразования строки даты из формата DD.MM.YYYY в формат YYYY-MM-DD
    function convertDate(dateStr) {
        // Разбиваем строку на части
        var parts = dateStr.split('.');
        // Возвращаем дату в формате YYYY-MM-DD
        return parts[2] + '-' + parts[1] + '-' + parts[0];
    }

    // Сортировка событий от самого нового к самому старому
    parsedEvents.sort(function(a, b) {
        var aDate = new Date(convertDate(a.dateTime.split(' в ')[0]));
        var bDate = new Date(convertDate(b.dateTime.split(' в ')[0]));
        // Для сортировки от новых к старым меняем местами b и a
        return bDate - aDate;
    });

    // Формирование итоговой строки с комментариями и причинами отмены, если они есть
    var result = parsedEvents.map(function(event) {
        var res = event.dateTime + " ";
        if (event.comment) {
            res += event.comment;
        }
        if (event.cancellationReason) {
            res += event.cancellationReason;
        }
        return res;
    }).join('\n');

    // Добавление оценок в конец списка событий
    if (scores) {
        result += "\n" + scores;
    }

    return result;
}

// function getEventDetailsForCandidate(candidateName) {
//     if (!candidateName) candidateName = 'Дементьев Евгений';

//     var now = new Date();
//     var thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
//     var thirtyDaysLater = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

//     var calendarIds = ['daria.rozhnova@sharp-dev.net', 'diana.berulava@sharp-dev.net', 'julia.kucherenko@sharp-dev.net', 'sedova@sharp-dev.net'];
//     var { allEvents, eventToCalendarMap } = fetchEventsForCalendars(calendarIds, thirtyDaysAgo, thirtyDaysLater);

//     var eventsDetailsForCandidate = getEventsDetails(allEvents, eventToCalendarMap, candidateName);
//     console.log(eventsDetailsForCandidate);
//     return eventsDetailsForCandidate;
// }


function getEventsDetails(allEvents, eventToCalendarMap, candidateName) {
    var eventsDetails = "";
    var addedEvents = new Set();

    allEvents.forEach(function(event) {
        var title = event.getTitle();
        var eventDate = event.getStartTime();
        var formattedDate = Utilities.formatDate(eventDate, "Europe/Moscow", "dd.MM.yyyy HH:mm");
        var eventKey = `${formattedDate} ${title}`;

        if (!addedEvents.has(eventKey) && title.toLowerCase().includes(candidateName.toLowerCase())) {
            addedEvents.add(eventKey);

            var attendeesSet = new Set();
            attendeesSet.add(eventToCalendarMap.get(event.getId()));
            event.getGuestList().forEach(guest => attendeesSet.add(guest.getEmail()));

            var attendees = Array.from(attendeesSet);
            attendees.sort();
            var attendeesStr = attendees.join(", ");

            var description = event.getDescription() || "No description provided";
            var cancellationReason = description.split('\n').find(line => line.trim().startsWith("Cancellation Reason:")) || "";
            if (cancellationReason) {
                cancellationReason = cancellationReason.substring("Cancellation Reason:".length).trim();
            }

            eventsDetails += `${formattedDate}\n${title}\nAttendees: ${attendeesStr}`;
            if (cancellationReason) {
                eventsDetails += `\nCancellation Reason: ${cancellationReason}`;
            }
            eventsDetails += "\n\n===============\n\n";
        }
    });

    return eventsDetails;
}


function fetchEventsForCalendars(calendarIds, thirtyDaysAgo, thirtyDaysLater) {
    var allEvents = [];
    var eventToCalendarMap = new Map();

    calendarIds.forEach(function(calendarId) {
        var calendar = CalendarApp.getCalendarById(calendarId);
        if (calendar) {
            var events = calendar.getEvents(thirtyDaysAgo, thirtyDaysLater);
            events.forEach(function(event) {
                var eventId = event.getId();
                if (!eventToCalendarMap.has(eventId)) {
                    allEvents.push(event);
                    eventToCalendarMap.set(eventId, calendarId);
                }
            });
        }
    });

    // Сортировка событий по дате и по ID, если даты совпадают
    allEvents.sort((a, b) => {
        var aDate = new Date(a.getStartTime() || a.getStartDate());
        var bDate = new Date(b.getStartTime() || b.getStartDate());
        var dateDiff = bDate - aDate; // Сортировка по дате в обратном порядке
        if (dateDiff !== 0) {
            return dateDiff;
        } else {
            var aId = a.getId();
            var bId = b.getId();
            return aId.localeCompare(bId);
        }
    });

    return { allEvents, eventToCalendarMap };
}


// function getEventDetailsForCandidateOld(candidateName) {
//   if (!candidateName) candidateName = 'Дементьев Евгений';
//     var now = new Date();
//     var thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
//     var thirtyDaysLater = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

//     var thirtyDaysAgoISOString = thirtyDaysAgo.toISOString();
//     var thirtyDaysLaterISOString = thirtyDaysLater.toISOString();

//     var calendarIds = ['daria.rozhnova@sharp-dev.net', 'diana.berulava@sharp-dev.net', 'julia.kucherenko@sharp-dev.net', 'sedova@sharp-dev.net'];
//     var token = PropertiesService.getScriptProperties().getProperty('OAUTH_TOKEN');

//     var options = {
//         headers: {
//             Authorization: 'Bearer ' + token
//         }
//     };

//     var allEventsForCandidate = [];
//     var uniqueEventsKeys = new Set();

//     calendarIds.forEach(function(calendarId) {
//         var pageToken;
//         do {
//             var url = "https://www.googleapis.com/calendar/v3/calendars/" + encodeURIComponent(calendarId) + 
//                       "/events?timeMin=" + thirtyDaysAgoISOString + "&timeMax=" + thirtyDaysLaterISOString + 
//                       (pageToken ? "&pageToken=" + pageToken : "");
//             var response = UrlFetchApp.fetch(url, options);
//             var data = JSON.parse(response.getContentText());

//             data.items.forEach(event => {
//                 if (!event.start) {
//                     return; // Skip events without a start date
//                 }
//                 var eventDate = event.start.dateTime || event.start.date;
//                 var formattedDate = Utilities.formatDate(new Date(eventDate), "Europe/Moscow", "dd.MM.yyyy HH:mm");
//                 var title = event.summary || "";
//                 var eventKey = `${formattedDate} ${title}`;

//                 if (!uniqueEventsKeys.has(eventKey) && title.toLowerCase().includes(candidateName.toLowerCase())) {
//                     uniqueEventsKeys.add(eventKey);
//                     allEventsForCandidate.push(event);
//                 }
//             });

//             pageToken = data.nextPageToken;
//         } while (pageToken);
//     });

//     //console.log(allEventsForCandidate);

//     // Sort events by date, and by ID if dates are the same
//     allEventsForCandidate.sort((a, b) => {
//         var aDate = new Date(a.start.dateTime || a.start.date);
//         var bDate = new Date(b.start.dateTime || b.start.date);
//         var dateDiff = bDate - aDate; // Сортировка по дате в обратном порядке
//         if (dateDiff !== 0) {
//             return dateDiff;
//         } else {
//             // Если даты событий одинаковы, сортируем по ID события
//             var aId = a.id; // Используем ID из объекта события
//             var bId = b.id;
//             // Сравнение строковых ID, если они строковые, или числовых, если они числовые
//             return aId.localeCompare(bId);
//         }
//     });

//     // Form event details string
//     var eventsDetails = allEventsForCandidate.map(event => {
//         var eventDate = new Date(event.start.dateTime || event.start.date);
//         var formattedDate = Utilities.formatDate(eventDate, "Europe/Moscow", "dd.MM.yyyy HH:mm");
//         var attendees = event.attendees ? event.attendees.map(a => a.email) : [];
//         attendees.sort(); // Сортировка списка гостей по алфавиту
//         var attendeesList = attendees.join(", ");
//         var description = event.description || "No description provided";
//         var title = event.summary || "";
//         var details = `${formattedDate}\n${title}\nAttendees: ${attendeesList}`;
//         var cancellationReason = description.split('\n').find(line => line.trim().startsWith("Cancellation Reason:")) || "";
//         if (cancellationReason) {
//             cancellationReason = cancellationReason.substring("Cancellation Reason:".length).trim();
//             details += `\nCancellation Reason: ${cancellationReason}`;
//         }

//         return details + "\n\n===============\n\n";
//     }).join('');

//     return eventsDetails;
// }

function updateCandidatesWithEvents() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var candidatesSheet = spreadsheet.getSheetByName("Candidates");
    var candidatesData = candidatesSheet.getDataRange().getValues();

    var now = new Date();
    var sixtyDaysAgo = new Date(now.getTime() - 60 * 24 * 60 * 60 * 1000);
    var thirtyDaysLater = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

    var calendarIds = ['daria.rozhnova@sharp-dev.net', 'danil.lobanov@sharp-dev.net', 'diana.berulava@sharp-dev.net', 'julia.kucherenko@sharp-dev.net', 'sedova@sharp-dev.net'];
    var { allEvents, eventToCalendarMap } = fetchEventsForCalendars(calendarIds, sixtyDaysAgo, thirtyDaysLater);

    if (allEvents.length === 0) return;

    var startRowIndex = candidatesSheet.getDataRange().getLastRow() - 50;
    if (startRowIndex === undefined) return;

    var nameColumnIndex = 0;
    var candidatesNamesAndRows = [];
    for (var i = startRowIndex; i < candidatesData.length; i++) {
        var name = candidatesData[i][nameColumnIndex];
        if (name) candidatesNamesAndRows.push({name: name, row: i + 1});
    }

    candidatesNamesAndRows.forEach(function(candidate) {
        var eventsDetailsForCandidate = getEventsDetails(allEvents, eventToCalendarMap, candidate.name);

        if (eventsDetailsForCandidate) {
            candidatesSheet.getRange(candidate.row, 67).setValue(eventsDetailsForCandidate);
            Logger.log("Обновлена информация для кандидата: " + candidate.name);
        }
    });
}



// function updateCandidatesWithEventsOld() {
//     var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//     var candidatesSheet = spreadsheet.getSheetByName("Candidates");
//     var candidatesData = candidatesSheet.getDataRange().getValues();

//     var now = new Date();
//     var thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
//     var thirtyDaysLater = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

//     var calendarIds = ['daria.rozhnova@sharp-dev.net', 'diana.berulava@sharp-dev.net', 'julia.kucherenko@sharp-dev.net', 'sedova@sharp-dev.net'];
//     var allEvents = [];

//     // Словарь для сохранения соответствия между событием и его календарем
//     var eventToCalendarMap = new Map();

//     calendarIds.forEach(function(calendarId) {
//         var calendar = CalendarApp.getCalendarById(calendarId);
//         if (calendar) {
//             var events = calendar.getEvents(thirtyDaysAgo, thirtyDaysLater);
//             events.forEach(function(event) {
//                 var eventId = event.getId();
//                 if (!eventToCalendarMap.has(eventId)) { // Проверяем, содержит ли map уже это событие
//                     allEvents.push(event);
//                     eventToCalendarMap.set(eventId, calendarId);
//                 } else {
//                     // Возможно, здесь стоит обновить данные в map, если это необходимо
//                     // Например, если вы хотите хранить календарь, который последний раз содержал событие:
//                     // eventToCalendarMap.set(eventId, calendarId);
//                 }
//             });
//         }
//     });


//     // Sort events by date, and by ID if dates are the same
//     allEvents.sort((a, b) => {
//         var aDate = new Date(a.getStartTime() || a.getStartDate());
//         var bDate = new Date(b.getStartTime() || b.getStartDate());
//         var dateDiff = bDate - aDate; // Сортировка по дате в обратном порядке
//         if (dateDiff !== 0) {
//             return dateDiff;
//         } else {
//             // Если даты событий одинаковы, сортируем по ID события
//             var aId = a.getId();
//             var bId = b.getId();
//             // Если ID представлены как строки, для их сравнения можно использовать localeCompare
//             return aId.localeCompare(bId);
//         }
//     });

//     if (allEvents.length === 0) return;

//     var startRowIndex = candidatesSheet.getDataRange().getLastRow() - 50;
//     Logger.log("Ищем со строки " + startRowIndex);

//     if (startRowIndex === undefined) return;

//     var nameColumnIndex = 0;
//     var candidatesNamesAndRows = [];
//     for (var i = startRowIndex; i < candidatesData.length; i++) {
//         var name = candidatesData[i][nameColumnIndex];
//         if (name) candidatesNamesAndRows.push({name: name, row: i + 1});
//     }

//     candidatesNamesAndRows.forEach(function(candidate) {
//         var eventsDetailsForCandidate = "";
//         var addedEvents = new Set();

//         allEvents.forEach(function(event) {
//             var title = event.getTitle();
//             var eventDate = event.getStartTime();
//             var formattedDate = Utilities.formatDate(eventDate, "Europe/Moscow", "dd.MM.yyyy HH:mm");
//             var eventKey = `${formattedDate} ${title}`;

//             if (!addedEvents.has(eventKey) && title.toLowerCase().includes(candidate.name.toLowerCase())) {
//                 addedEvents.add(eventKey);

//                 // Используем Set для сбора и автоматического удаления дубликатов email'ов участников
//                 var attendeesSet = new Set();

//                 // Добавляем calendarId как первого гостя
//                 attendeesSet.add(eventToCalendarMap.get(event.getId()));

//                 // Добавляем email всех гостей, автоматически удаляя дубликаты
//                 event.getGuestList().forEach(guest => attendeesSet.add(guest.getEmail()));

//                 // Преобразуем Set в массив для сортировки
//                 var attendees = Array.from(attendeesSet);
//                 attendees.sort(); // Сортировка списка гостей по алфавиту
//                 var attendeesStr = attendees.join(", ");

//                 // Добавляем детали события
//                 var description = event.getDescription() || "No description provided";
//                 var cancellationReason = description.split('\n').find(line => line.trim().startsWith("Cancellation Reason:")) || "";
//                 if (cancellationReason) {
//                     cancellationReason = cancellationReason.substring("Cancellation Reason:".length).trim();
//                 }

//                 eventsDetailsForCandidate += `${formattedDate}\n${title}\nAttendees: ${attendeesStr}`;
//                 if (cancellationReason) {
//                     eventsDetailsForCandidate += `\nCancellation Reason: ${cancellationReason}`;
//                 }
//                 eventsDetailsForCandidate += "\n\n===============\n\n";
//             }
//         });

//         if (eventsDetailsForCandidate) {
//             candidatesSheet.getRange(candidate.row, 67).setValue(eventsDetailsForCandidate);
//             Logger.log("Обновлена информация для кандидата: " + candidate.name);
//         }
//     });

// }


function getLastModifiedDate(docId) {
    var file = DriveApp.getFileById(docId);
    return file.getLastUpdated();
}

function updateInterviewConclusions() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var candidatesSheet = sheet.getSheetByName("Candidates");
    var recentModSheet = sheet.getSheetByName("CandidatesRecentModifiedFilesData");
    var data = candidatesSheet.getDataRange().getValues();
    var recentData = recentModSheet.getDataRange().getValues();

    var recentModData = {};
    recentData.forEach(function(row, index) {
        if (index > 0) { // Пропускаем заголовок
            recentModData[row[1]] = { // ID файла как ключ
                fileName: row[0],
                lastModifiedDate: new Date(row[2]),
                fileLink: row[3],
                rowIndex: index + 1 // Сохраняем индекс строки для использования в листе
            };
        }
    });

    for (var i = data.length; i > data.length - 150; i--) {
        var docId = docidfromurl(i + 1, 36);
        if (!docId) continue;

        var fileInfo = recentModData[docId];
        if (!fileInfo) continue; // Если нет информации о файле, пропускаем

        var dateCell = candidatesSheet.getRange(i + 1, 69).getValue();
        var lastColumnUpdateDate = dateCell ? new Date(dateCell) : null;
        var currentDate = new Date();

        Logger.log("Обработка файла: " + fileInfo.fileName + " (ID: " + docId + ")");
        Logger.log("Ссылка на файл: " + fileInfo.fileLink);
        Logger.log("Дата последнего изменения файла: " + fileInfo.lastModifiedDate);
        Logger.log("Дата последнего изменения в колонке: " + (lastColumnUpdateDate ? lastColumnUpdateDate : "Нет данных"));

        // Сравниваем даты
        if (!lastColumnUpdateDate || fileInfo.lastModifiedDate > lastColumnUpdateDate) {
            Logger.log("Документ был обновлен после последнего изменения в колонке.");

            var interviewData = parseInterviewNew(docId);
            var conclusion = parseConclusion(i + 1, 36); // Получение вывода

            candidatesSheet.getRange(i + 1, 9).setValue(interviewData.text);         // Колонка I
            candidatesSheet.getRange(i + 1, 10).setValue(interviewData.technologyText);     // Колонка J
            candidatesSheet.getRange(i + 1, 11).setValue(interviewData.englishText);        // Колонка K
            candidatesSheet.getRange(i + 1, 14).setValue(interviewData.educationText);      // Колонка N

            // Предполагаем, что customFields сохраняют информацию для кастомных полей от BC до BL
            const startColumnIndex = 55; // Начиная с колонки BC
            const columns = ['BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL']; // Соответствие ключевым полям
            columns.forEach((column, index) => {
                if (interviewData.customText[column]) {
                    candidatesSheet.getRange(i + 1, startColumnIndex + index).setValue(interviewData.customText[column]);
                }
            });

            candidatesSheet.getRange(i + 1, 68).setValue(conclusion); // Обновление значения в колонке
            candidatesSheet.getRange(i + 1, 69).setValue(currentDate);
        } else {
            Logger.log("Файл не был изменен после последнего записанного изменения в колонке. Пропуск обработки.");
        }
    }
}


// function updateInterviewConclusions() {
//     var sheet = SpreadsheetApp.getActiveSpreadsheet();
//     var candidatesSheet = sheet.getSheetByName("Candidates");
//     var data = candidatesSheet.getDataRange().getValues();

//     // Получаем ключевые слова из первой строки для колонок BC до BL
//     //var keywords = candidatesSheet.getRange(1, 55, 1, 12).getValues()[0]; // Колонки от BC (индекс 55) до BL (индекс 66)

//     for (var i = data.length; i > data.length - 50; i--) {
//         var docId = docidfromurl(i+1, 36);
//         console.log(docId);
//         if(!docId) continue;
//         //var docLink = data[i][35]; // Предполагаем, что ссылка на документ находится в колонке AJ
//         //var docId = extractDocIdFromUrl(docLink); // Функция для извлечения ID документа из URL
//         //var lastUpdated = getLastModifiedDate(docId);

//         //Logger.log("Последнее обновление документа: " + lastUpdated);

//         var dateCell = candidatesSheet.getRange(i + 1, 69).getValue();
//         var lastColumnUpdateDate = new Date(dateCell);
//         var currentDate = new Date();

//         Logger.log("Последнее обновление колонки: " + lastColumnUpdateDate);

//         // Сравнение даты последнего обновления документа и колонки
//         if (!lastColumnUpdateDate || !dateCell) { // || lastUpdated > lastColumnUpdateDate
//             Logger.log("Документ был обновлен после последнего изменения в колонке, продолжаем обработку.");

//             var interviewData = parseInterviewNew(docId);
//             var conclusion = parseConclusion(i + 1, 36); // Получение вывода

//             candidatesSheet.getRange(i + 1, 9).setValue(interviewData.text);         // Колонка I
//             candidatesSheet.getRange(i + 1, 10).setValue(interviewData.technologyText);     // Колонка J
//             candidatesSheet.getRange(i + 1, 11).setValue(interviewData.englishText);        // Колонка K
//             candidatesSheet.getRange(i + 1, 14).setValue(interviewData.educationText);      // Колонка N

//             // Предполагаем, что customFields сохраняют информацию для кастомных полей от BC до BL
//             const startColumnIndex = 55; // Начиная с колонки BC
//             const columns = ['BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL']; // Соответствие ключевым полям
//             columns.forEach((column, index) => {
//                 if (interviewData.customText[column]) {
//                     candidatesSheet.getRange(i + 1, startColumnIndex + index).setValue(interviewData.customText[column]);
//                 }
//             });

//             candidatesSheet.getRange(i + 1, 68).setValue(conclusion); // Обновление значения в колонке
//             candidatesSheet.getRange(i + 1, 69).setValue(currentDate);
//         } else {
//             Logger.log("Документ не обновлялся после последнего изменения в колонке. Пропуск обработки.");
//         }
//     }
// }


// function updateInterviewConclusions() {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet();
//   var candidatesSheet = sheet.getSheetByName("Candidates");
//   var data = candidatesSheet.getDataRange().getValues();

//   // Определить начальный индекс для обработки последних 50 строк
//   var startRow = Math.max(data.length - 50, 1); // Убедимся, что индекс не выходит за пределы возможного (если строк меньше 50)

//   // Пройти по последним 50 строкам данных
//   for (var i = startRow - 1; i < data.length; i++) {
//     var conclusion = parseConclusion(i + 1, 36); // Получить вывод, предполагается, что функция возвращает текст

//     // Вывести информацию для отладки
//     console.log("Row: " + (i + 1));
//     console.log("Conclusion: " + conclusion);

//     // Обновить значение в колонке BP (индекс 69)
//     candidatesSheet.getRange(i + 1, 68).setValue(conclusion);
//   }
// }


// function updateInterviewConclusions() {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet();
//   var candidatesSheet = sheet.getSheetByName("Candidates");
//   var data = candidatesSheet.getDataRange().getValues();

//   var currentDate = new Date();
//   var sevenDaysAgo = new Date();
//   var futureDate = new Date();
//   sevenDaysAgo.setDate(currentDate.getDate() - 7);
//   futureDate.setDate(currentDate.getDate() + 14);


//   // Пройти по всем строкам и проверить даты в колонке AG (индекс 32, так как отсчет начинается с 0)
//   for (var i = 0; i < data.length; i++) {
//     var rowDate = new Date(data[i][32]); // Дата в строке
//     if (rowDate >= sevenDaysAgo && rowDate <= futureDate) {
//       // Если дата находится в нужном диапазоне, обработать строку
//       // var docLink = data[i][35]; // Ссылка на интервью в колонке AJ (индекс 35)
//       var conclusion = parseConclusion(i + 1, 36); // Получить вывод, предполагается, что функция возвращает текст

//       var row = i+1;
//       console.log(row);
//       console.log(conclusion);
//       // Обновить значение в колонке BP (индекс 69)
//       candidatesSheet.getRange(i + 1, 68).setValue(conclusion);
//     }
//   }
// }


function extractFileId(url) {
    var match = url.match(/[-\w]{25,}/);
    return match ? match[0] : null;
}

function parseConclusion(row, column) {
    Logger.log(row + " - " + column);
    var docId = docidfromurl(row, column);
    if (!docId) return;
    Logger.log(docId);
    var token = PropertiesService.getScriptProperties().getProperty('OAUTH_TOKEN');
    var url = "https://docs.googleapis.com/v1/documents/" + docId; // URL для запроса
    Logger.log(url);
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
    var allText = '';

    Logger.log(structuralElements.length);

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
                    allText += textContent;
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

                            // проверка: является ли элемент текстовым run и содержит ли контент
                            if (textRun.textRun && textRun.textRun.content) {
                                var textContent = textRun.textRun.content;
                                text += textContent; // добавляем текст
                                Logger.log("Найден текст: " + textContent);
                            } else {
                                Logger.log("Пропущен не-текстовый элемент или пустой контент");
                            }
                        }
                    }

                    if (cellElement.inlineObjectElement) { // если элемент - inline object
                        var inlineObjectElement = cellElement.inlineObjectElement;
                        var inlineObjectId = inlineObjectElement.inlineObjectId;
                        Logger.log("Найден inline object с ID: " + inlineObjectId);
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

    if(text === "") text = allText;

    return text + total;

}

function checkStringForMatch(str) {
    var skillsMap = {
        "general": "Тех. собеседование",
        "netremove": ".NET (убрать",
        "netnew": ".Net (new, WIP)",
        "nodejs": "NodeJS",
        "vanillajs" : "VanillaJS",
        "react": "React",
        "react-native": "React-Native",
        "ios": "iOS Native (Swift)",
        "android": "Android Native (Kotlin/Java)",
        "flutter": "Flutter",
        "angular": "Angular",
        "python": "Python",
        "django": "Django",
        "phpnew": "PHP (new)",
        "phpold": "PHP (old)",
        "laravel": "Laravel",
        "wordpress": "WordPress",
        "go language": "Go language",
        "testing": "Testing",
        "analysis": "Analysis",
        "java": "Java",
        "multithreading": "Multithreading",
        "oracle": "Oracle",
        "databases": "Базы данных",
        "apparchitecture": "Архитектура приложений",
        "devopsazure": "Devops Azure",
        "devopsgcp": "Devops Google Cloud Platform",
        "devopsaws": "Devops AWS",
        "devops": "Devops General",
        "smart_contract": "Smart contract (Solidity, Rust)",
        "datascience" : "Data science",
        "htmlcoding" : "Вёрстка"

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
    if(!url) return;
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
    if (fields.length < 2) return fields.allText;
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



function parseInterviewNew(docId) {
    if (!docId) {
        docId = "1oyGaEviqP9WfI_5W7_c0942TIiSbSBt0OCSMT8HQ3L0"; // идентификатор документа по умолчанию
    }

    // Если данных в кэше нет, выполняем запрос и обработку документа
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var candidatesSheet = spreadsheet.getSheetByName("Candidates");
    var firstRowData = candidatesSheet.getRange("BC1:BL1").getValues()[0]; // Извлекаем данные из первой строки от столбца BC до BL


    var fields = {
        customText: {}
    };

    const columns = ['BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL'];

    // Обработка документа для заполнения остальных полей (примерная логика)
    var document = fetchDocumentContent(docId);

    var lastTable = null; // переменная для хранения последней таблицы
    var structuralElements = document.body.content; // массив structural elements в документе
    var text = '';
    var allText = '';
    var technologyText = "";
    var englishText = "";
    var educationText = "";
    var current = "";
    var customKeyword = "про Строителя";

    var keywords = {
        "english": "Английский",
        "experience": "Коммерческий опыт",
        "education": "Высшее образование",
        "remote": "Опыт удаленной работы",
        "computer": "Есть ли хороший домашний компьютер",
        "quiz": "Quiz",
        "softskills": "Софт-скиллы",
        "tracker": "ТРЕКЕР и как обычно отслеживает время",
        "technology": "Технологии:",
        "location": "ЛОКАЦИЯ\\/РЕЛОКАЦИЯ",
        "legal": "ИП\\/САМОЗАНЯТЫЙ",
        "business": "Есть ли у вас сейчас свой бизнес",
        "daysoff": "Планируете ли вы отпуск",
        "latestproject": "Последний проект",
    };

    columns.forEach((column, index) => {
        keywords[column] = firstRowData[index]; // Map column letters to keywords
        fields.customText[column] = ""; // Initialize text storage for each keyword
    });


    // if (customKeyword) {
    //     keywords["custom"] = customKeyword;
    // }

    console.log(keywords);

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
                    allText += textContent;
                    //Logger.log("Найден текст: " + textContent);
                    if(!customKeyword && !stop) {
                        var stop = textContent.match(/(Самая сложная задача|Culture fit)\s*/); // Как нашли в тексте Самая сложная задача или Culture fit - останавливаем поиск
                    }
                    if(customKeyword && !customStop) {
                        var customStop = textContent.match(new RegExp(customKeyword + '\\s*'));  // Как нашли в тексте custom поле - останавливаем поиск
                    }

                    var matches = {};
                    for (var key in keywords) {
                        var regex = new RegExp(keywords[key], "i"); // "i" указывает на регистронезависимый поиск
                        matches[key] = textContent.match(regex);
                    }


                    for (var key in matches) {
                        if (matches[key]) {
                            text += matches[key][0];
                            //Logger.log('найден ключ:' + matches[key][0]);
                            current = key;
                        }
                    }


                }
            }
        }
        var lastTable = element.table; // запомнить его как последнюю таблицу

        if (lastTable) { // если нашли хотя бы одну таблицу
            //Logger.log("найдена таблица!");
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
                                    //Logger.log("В таблице найден текст");
                                    //Logger.log(textContent);

                                    var serviceField = textContent.match(/(Оценка|Ответ|ЗП)\s*/);
                                    if (!serviceField) {
                                        textContent = textContent.replace(/: \n/g, " ");
                                        //Logger.log(textContent + " - текст удовлетворяет условиям");
                                        if (current && current !== "technology") text += '\n' + textContent;   //  && current !== "education" && current !== "english" && !current.startsWith("B")

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
                                            case "BC":
                                                fields.customText['BC'] += textContent;
                                                break;
                                            case "BD":
                                                fields.customText['BD'] += textContent;
                                                break;
                                            case "BE":
                                                fields.customText['BE'] += textContent;
                                                break;
                                            case "BF":
                                                fields.customText['BF'] += textContent;
                                                break;
                                            case "BG":
                                                fields.customText['BG'] += textContent;
                                                break;
                                            case "BH":
                                                fields.customText['BH'] += textContent;
                                                break;
                                            case "BI":
                                                fields.customText['BI'] += textContent;
                                                break;
                                            case "BJ":
                                                fields.customText['BJ'] += textContent;
                                                break;
                                            case "BK":
                                                fields.customText['BK'] += textContent;
                                                break;
                                            case "BL":
                                                fields.customText['BL'] += textContent;
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
    //fields.customText = replaceLinebreaks(customText);
    //fields.customText[]
    if (fields.length < 5) fields.allText = replaceLinebreaks(allText);

    console.log(fields);

    return fields; // Возвращаем собранные и обработанные данные
}



function parseInterview(docId,customKeyword) {
    // input - docId документа, который нужно распарсить
    if(!docId) var docId = "171cigxFBFqEBoa88kQtpIlryaVC1wiO16HDuZVjc9Gg"; // идентификатор документа

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
    var allText = '';
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
                    allText += textContent;
                    Logger.log("Найден текст: " + textContent);
                    if(!customKeyword && !stop) {
                        var stop = textContent.match(/(Дополнительно|Culture fit)\s*/); // Как нашли в тексте Самая сложная задача или Culture fit - останавливаем поиск
                    }
                    if(customKeyword && !customStop) {
                        var customStop = textContent.match(new RegExp(customKeyword + '\\s*'));  // Как нашли в тексте custom поле - останавливаем поиск
                    }
                    var keywords = {
                        "english": "Английский",
                        "experience": "Коммерческий опыт",
                        "education": "Высшее образование",
                        "remote": "Опыт удаленной работы",
                        "computer": "Есть ли хороший домашний компьютер",
                        "quiz": "Quiz",
                        "softskills": "Софт-скиллы",
                        "tracker": "ТРЕКЕР и как обычно отслеживает время",
                        "technology": "Технологии",
                        "location": "ЛОКАЦИЯ\\/РЕЛОКАЦИЯ",
                        "legal": "ИП\\/САМОЗАНЯТЫЙ",
                        "business": "Есть ли у вас сейчас свой бизнес",
                        "daysoff": "Планируете ли вы отпуск",
                        "latestproject": "Последний проект",
                    };

                    if (customKeyword) {
                        keywords["custom"] = customKeyword;
                    }

                    var matches = {};
                    for (var key in keywords) {
                        var regex = new RegExp(keywords[key], "i"); // "i" указывает на регистронезависимый поиск
                        matches[key] = textContent.match(regex);
                    }

                    for (var key in matches) {
                        if (matches[key]) {
                            text += matches[key][0];
                            Logger.log('найден ключ:' + matches[key][0]);
                            current = key;
                        }
                    }

                }
            }
        }
        var lastTable = element.table; // запомнить его как последнюю таблицу

        if (lastTable) { // если нашли хотя бы одну таблицу
            Logger.log("найдена таблица!");
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
                                    Logger.log("В таблице найден текст");
                                    Logger.log(textContent);
                                    var serviceField = textContent.match(/(Оценка|Ответ|ЗП)\s*/);
                                    if (!serviceField) {
                                        textContent = textContent.replace(/: \n/g, " ");
                                        Logger.log(textContent + " - текст удовлетворяет условиям");
                                        if (current && current !== "technology") text += textContent;

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
    fields.allText = replaceLinebreaks(allText);
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


function getLinkFromRichText(cellAddress) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getRange(cellAddress);
    var richTextValue = range.getRichTextValue();
    var url = richTextValue.getLinkUrl();
    return url;
}


function fillFormulas(row) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    if(row) {
        var lastRow = row;
    } else {
        var row = 2;
        var lastRow = sheet.getLastRow();
    }

    for (var i = row; i <= lastRow; i++) {
        var cellA = "A" + i;
        var cellG = "G" + i;
        var cellAH = "AH" + i;
        var cellAJ = "AJ" + i;
        var cellL = "L" + i;

        var formula = '=' + cellAH + ' & " мск" & CHAR(10) & ' + cellA + ' & " (" & ' + cellG + ' & ")\nФайл: " & getLinkFromRichText("' + cellAJ + '") & " \nC/V: " & getLinkFromRichText("' + cellL + '")';

        var range = sheet.getRange("BN" + i); // Замените "Z" на колонку, где вы хотите поместить формулу
        range.setFormula(formula);
    }
}

function findDuplicates() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Кандидаты сорс/скрин");
    var candidatesSheet = ss.getSheetByName("Candidates");
    var tanyaSheet = ss.getSheetByName("Кандидаты Таня"); // Лист для сверки
    var sourceData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn()).getValues();
    var candidatesData = candidatesSheet.getRange(1, 1, candidatesSheet.getLastRow(), candidatesSheet.getLastColumn()).getValues();
    var tanyaData = tanyaSheet.getRange("A:A").getValues(); // Получение данных из колонки A листа "Кандидаты Таня"
    var reportColumnIndex = null;
    var sourceIdColumnIndex = null;
    var candidatesIdColumnIndex = null;

    // Находим индекс колонки для отчета и ID в исходном листе
    var sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    sourceHeaders.forEach(function(header, index) {
        if (header === "Дубликаты") {
            reportColumnIndex = index + 1;
        } else if (header === "ID") { // Предполагаем, что заголовок столбца с ID - это "ID"
            sourceIdColumnIndex = index;
        }
    });

    // Находим индекс колонки с ID в листе candidates
    var candidatesHeaders = candidatesSheet.getRange(1, 1, 1, candidatesSheet.getLastColumn()).getValues()[0];
    candidatesHeaders.forEach(function(header, index) {
        if (header === "ID") {
            candidatesIdColumnIndex = index;
        }
    });

    if (reportColumnIndex === null) {
        Logger.log("Колонка 'Дубликаты' не найдена.");
        return;
    }

    if (sourceIdColumnIndex === null || candidatesIdColumnIndex === null) {
        Logger.log("Колонка 'ID' не найдена в одном из листов.");
        return;
    }

    // Подготовка списка имен, Telegram и ID для поиска
    var namesAndTelegrams = sourceData.map(function(row, index) {
        return {
            russianName: row[2].toString().toLowerCase().replace(/[^a-zа-яё0-9\s]/gi, '').trim().split(/\s+/),
            englishName: row[3].toString().toLowerCase().replace(/[^a-zа-яё0-9\s]/gi, '').trim().split(/\s+/),
            telegram: row[12].toString(),
            id: row[sourceIdColumnIndex].toString(),
            rowIndex: index + 2
        };
    });


    namesAndTelegrams.forEach(function(item) {
        // Пропускаем итерацию, если ID пуст
        if (!item.id.trim()) {
            return; // Продолжаем с следующего элемента массива
        }

        // Пропускаем итерацию, если все проверяемые значения пусты
        if (!item.russianName.join('').trim() && !item.englishName.join('').trim() && !item.telegram.trim()) {
            return; // Продолжаем с следующего элемента массива
        }

        var matches = []; // Для совпадений с "Candidates"
        var tanyaMatches = []; // Для совпадений с "Кандидаты Таня"
        var matchString = "";

        // Поиск совпадений по имени и Telegram, исключая строки с тем же ID, в "Candidates"
        candidatesData.forEach(function(candidate, index) {
            if (candidate[candidatesIdColumnIndex] === item.id) return; // Пропускаем строки с тем же ID

            var candidateName = candidate[0].toLowerCase();
            if (item.russianName.every(function(namePart) { return candidateName.includes(namePart.toLowerCase()); }) ||
                item.englishName.every(function(namePart) { return candidateName.includes(namePart.toLowerCase()); })) {
                matches.push(index + 1);
            }

            if (item.telegram && candidate.includes(item.telegram)) {
                matches.push(index + 1);
            }
        });

        // Поиск совпадений по имени в "Кандидаты Таня" с той же логикой
        tanyaData.forEach(function(row, index) {
            if (!row[0] || row[0].toString().trim() === '') return; // Пропускаем пустые строки

            // Разделяем имя из "Кандидаты Таня" на слова, оставляя кириллические символы
            var tanyaNameParts = row[0].toString().toLowerCase().replace(/[^a-zа-яё0-9\s]/gi, '').trim().split(/\s+/);
            // console.log(tanyaNameParts);

            // Проверяем, что каждая часть имени и фамилии кандидата присутствует в имени из "Кандидаты Таня"
            var russianNameMatch = item.russianName.every(function(namePart) {
                // console.log(namePart);
                return tanyaNameParts.includes(namePart.toLowerCase().replace(/[^a-zа-яё0-9\s]/gi, ''));
            });

            var englishNameMatch = item.englishName.every(function(namePart) {
                return tanyaNameParts.includes(namePart.toLowerCase().replace(/[^a-zа-яё0-9\s]/gi, ''));
            });

            if (russianNameMatch || englishNameMatch) {
                tanyaMatches.push(index + 1); // Сохраняем номер строки для совпадений
            }
        });

        // Формирование строки отчета с учетом результатов поиска
        matchString += matches.length > 0 ? matches.length + " совпадений в 'Candidates'. Строки: " + matches.join(", ") : "Candidates - OK. ";
        matchString += tanyaMatches.length > 0 ? "; " + tanyaMatches.length + " совпадений в 'Кандидаты Таня'. Строки: " + tanyaMatches.join(", ") : "; Кандидаты Таня - ОК.";

        // Запись отчета в лист
        sourceSheet.getRange(item.rowIndex, reportColumnIndex).setValue(matchString);
    });

    Logger.log("Поиск дубликатов завершен.");
}


function updateCandidatesStatusAndNotes() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName("Кандидаты сорс/скрин");
    var candidatesSheet = ss.getSheetByName("Candidates");
    var lastRowSource = sourceSheet.getLastRow();
    var lastRowCandidates = candidatesSheet.getLastRow();

    // Получаем диапазоны данных
    var sourceData = sourceSheet.getRange(2, 1, lastRowSource, sourceSheet.getLastColumn()).getValues();
    var candidatesData = candidatesSheet.getRange(2, 1, lastRowCandidates, candidatesSheet.getLastColumn()).getValues();

    // Получаем индексы необходимых колонок
    var sourceStatusColumnIndex = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].indexOf("Candidates status") + 1;
    var idColumnIndexSource = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0].indexOf("ID") + 1;
    var idColumnIndexCandidates = candidatesSheet.getRange(1, 1, 1, candidatesSheet.getLastColumn()).getValues()[0].indexOf("ID") + 1;
    var interviewDateIndex = candidatesSheet.getRange(1, 1, 1, candidatesSheet.getLastColumn()).getValues()[0].indexOf("Interview Date") + 1;

    sourceData.forEach(function(row, index) {
        var id = row[idColumnIndexSource - 1];
        // Проверяем, заполнен ли ID
        if (!id) {
            // Если ID не заполнен, пропускаем текущую итерацию
            return;
        }

        var statusCell = sourceSheet.getRange(index + 2, sourceStatusColumnIndex);
        var status = row[sourceStatusColumnIndex - 1];

        // Поиск соответствующего ID в "Candidates"
        for (var i = 0; i < candidatesData.length; i++) {
            var candidate = candidatesData[i];
            if (candidate[idColumnIndexCandidates - 1] === id) {
                var interviewDate = candidate[interviewDateIndex - 1];
                var currentDate = new Date();
                var oneWeekAgo = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - 7);
                var halfYearAgo = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - 180);

                if (!status || (interviewDate && interviewDate >= halfYearAgo && interviewDate <= currentDate)) {
                    // Если статус пуст или дата интервью находится в допустимом диапазоне (не старше полугода)
                    var statusValue = candidate[2]; // Статус из колонки C
                    var finalComment = candidate[7]; // Финальный коммент из колонки H
                    var comments = candidate[37]; // Комментарии рекрутера и техлида из колонки AL

                    // Формируем текст для заметок
                    var notesText = statusValue + "\n\n" + finalComment + "\n\n" + comments;
                    var link = ss.getUrl() + "#gid=" + candidatesSheet.getSheetId() + "&range=" + (i + 2) + ":" + (i + 2);
                    var statusWithLink = '=HYPERLINK("' + link + '"; "' + statusValue + '")';
                    statusCell.setFormula(statusWithLink); // Обновляем статус с гиперссылкой
                    statusCell.setNote(notesText); // Обновляем заметки
                }
                break; // Выходим из цикла после обновления
            }
        }
    });

}



function insertFormulaForEmptyID() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Кандидаты сорс/скрин"); // Замените на нужное имя листа
    var lastRow = sheet.getLastRow();
    var dataRange = sheet.getRange(2, 1, lastRow, sheet.getLastColumn()); // Получаем данные начиная со 2-й строки, чтобы исключить заголовок
    var values = dataRange.getValues();

    // Перебираем строки и проверяем условия
    for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var name = row[0]; // Предполагаем, что имя находится в первой колонке
        var id = row[28]; // Предполагаем, что ID находится во второй колонке, корректируйте индекс в соответствии с расположением вашей колонки ID

        // Проверяем, не пустое ли имя и пустой ли ID
        if (name && !id) {
            // Формируем формулу с учетом текущей строки
            var formula = '=IFERROR(IFERROR(MID(E' + (i + 2) + '; SEARCH("/d/"; E' + (i + 2) + ') + 3; SEARCH("/edit"; E' + (i + 2) + ') - SEARCH("/d/"; E' + (i + 2) + ') - 3); MID(E' + (i + 2) + '; SEARCH("id="; E' + (i + 2) + ') + 3; LEN(E' + (i + 2) + ') - SEARCH("id="; E' + (i + 2) + ') - 2)); "")';
            // Вставляем формулу в ячейку ID текущей строки
            sheet.getRange(i + 2, 29).setFormula(formula); // Помните, что индексы в Google Sheets начинаются с 1, а не с 0
        }
    }
}


function fetchDocumentContent(docId) {
    var token = PropertiesService.getScriptProperties().getProperty('OAUTH_TOKEN');
    var url = "https://docs.googleapis.com/v1/documents/" + docId;
    var options = {
        "method": "get",
        "headers": {
            "Authorization": "Bearer " + token
        },
        'fields': 'documentStyle,updateTime'
    };
    var response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText()); // Возвращаем объект документа
}


function updateBenchStatus() {
    var benchSheetName = "Bench+";  // Название листа Bench+
    var workloadFileId = "16aHQE2D9RBC-GdbIxVjQtDzktd-nMwW-J5fhl5kzN8Y"; // ID документа Workload 2025

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var benchSheet = ss.getSheetByName(benchSheetName);
    if (!benchSheet) {
        Logger.log("Лист Bench+ не найден");
        return;
    }

    var benchData = benchSheet.getDataRange().getValues();
    var headers = benchData[0];

    var nameColIndex = 0; // "Фамилия Имя" всегда в первой колонке
    var workloadColIndex = headers.indexOf("Workload");
    var availableColIndex = headers.indexOf("Available?");
    var locationColIndex = headers.indexOf("Location");
    var stackColIndex = headers.indexOf("Stack");
    var gradeColIndex = headers.indexOf("Grade");

    if (workloadColIndex === -1 || availableColIndex === -1 || locationColIndex === -1 || stackColIndex === -1 || gradeColIndex === -1) {
        Logger.log("Одна или несколько колонок не найдены!");
        return;
    }

    // Открываем Workload 2025 и берём **крайний левый лист**
    var workloadSS = SpreadsheetApp.openById(workloadFileId);
    var workloadSheet = workloadSS.getSheets()[0];
    var workloadData = workloadSheet.getDataRange().getValues();

    // Получаем список имен из колонки D
    var workloadNames = workloadData.slice(1).map(row => row[3] ? row[3].trim() : "");
    var workloadSet = new Set(workloadNames); // Для быстрого поиска

    // Списки для отчета
    var addToWorkload = [["Location", "Добавить в Workload"]];
    var removeFromWorkload = [["Location", "Удалить из Workload"]];

    // Обновляем колонку "Workload" в Bench+
    var updated = false;
    for (var i = 1; i < benchData.length; i++) {
        var name = benchData[i][nameColIndex] ? benchData[i][nameColIndex].trim() : "";

        if (!name) continue; // Пропускаем пустые строки
        if (name.includes("Вышли к нам")) break; // Прерываем выполнение, если нашли "Вышли к нам"

        var location = benchData[i][locationColIndex] ? benchData[i][locationColIndex].trim() : "";
        var stack = benchData[i][stackColIndex] ? benchData[i][stackColIndex].trim() : "";
        var grade = benchData[i][gradeColIndex] ? benchData[i][gradeColIndex].trim() : "";
        var available = benchData[i][availableColIndex] === true; // TRUE/FALSE
        var isPresent = [...workloadSet].some(wName => wName.startsWith(name));

        var formattedName = `${name} (${stack} ${grade})`;

        if (available && !isPresent) {
            addToWorkload.push([location, formattedName]); // Добавляем в "Добавить в Workload"
        }

        if (!available && isPresent) {
            removeFromWorkload.push([location, formattedName]); // Добавляем в "Удалить из Workload"
        }

        var currentValue = benchData[i][workloadColIndex];
        var newValue = isPresent ? "✅" : "";

        if (currentValue !== newValue) {
            benchSheet.getRange(i + 1, workloadColIndex + 1).setValue(newValue);
            updated = true;
        }
    }

    // Публикуем отчет
    var reportSheet = ss.getSheetByName("Workload Report") || ss.insertSheet("Workload Report");
    reportSheet.clear(); // Очищаем предыдущий отчет
    reportSheet.getRange(1, 1, addToWorkload.length, addToWorkload[0].length).setValues(addToWorkload);
    reportSheet.getRange(addToWorkload.length + 2, 1, removeFromWorkload.length, removeFromWorkload[0].length).setValues(removeFromWorkload);

    if (updated) {
        Logger.log("Обновление завершено!");
    } else {
        Logger.log("Изменений нет.");
    }
}



