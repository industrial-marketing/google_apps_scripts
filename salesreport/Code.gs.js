function saveOAuthToken() {
    var token = ScriptApp.getOAuthToken();
    PropertiesService.getScriptProperties().setProperty('OAUTH_TOKEN', token);
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
        .addItem('Save OAuth Token', 'saveOAuthToken')
        .addItem('Generate SALES report', 'generateSalesReportCheckOAuth')
        .addItem('Update scrum files data for current week in "Scrum files for current week"', 'gatherDataInCurrentSheetCheckOAuth')
        .addItem('Update developers competences in "Competences"', 'copyDataToCompetencesSheetCheckOAuth')
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

function copyDataToCompetencesSheetCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию generateSalesReport()
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

var monthNamesShort = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

function generateSalesReport() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reportSheet = ss.getSheetByName('SALES report');

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

    const competencesSheetId = "1dblvYXs4QfW2WbZGlU489I6Yy9o0xc6s_33w96DdXlA";
    const competencesSpreadsheet = SpreadsheetApp.openById(competencesSheetId);
    const competencesSheet = competencesSpreadsheet.getSheetByName('Компетенции Dev');

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



    let developers = getDevelopers(workloadSheet);
    const { competenceDevelopers, competences } = getCompetences(competencesSheet, developers);

    // Initialize report
    reportSheet.clearContents();
    reportSheet.getRange('B3').setValue(`SALES report for ${currentWeekMondayString} - ${currentWeekSundayString}`).setFontSize(20);

    // Initialize the header row
    reportSheet.getRange('B5').setValue('Developer').setVerticalAlignment("middle");
    reportSheet.getRange('C5').setValue('English').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('D5').setValue('Training').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('E5').setValue('Sales').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('F5').setValue('Allocation').setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('G5').setValue('Resume Link').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('H5').setValue('Stack').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('I5').setValue('Extra stack').setTextRotation(90).setBackground("#ffffff").setVerticalAlignment("middle").setHorizontalAlignment("center");
    reportSheet.getRange('J5').setValue('Обучаемость\nСтрессоустойчивость\nРабота в команде\nРабота с клиентом\nНавыки самопрезентации\nГибкость мышления').setTextRotation(90).setBackground("#ffffff").setHorizontalAlignment("center").setVerticalAlignment("middle");

    let column = 11;
    for (let competence of competences) {
        reportSheet.getRange(5, column).setValue(competence).setVerticalAlignment("middle").setHorizontalAlignment("center").setTextRotation(90).setBackground("#cccccc");
        reportSheet.setColumnWidth(column, 20);
        column++;
    }


    // Initialize the data rows
    let row = 6;
    for (let developer of developers) {
        let developerName = developer.name.split("(")[0].trim(); // Remove everything after the "(" and trim spaces

        // Check if the key exists in the object
        //if (competenceDevelopers.hasOwnProperty(developerName)) {
        let developerData = competenceDevelopers[developerName];
        let englishLevel = getDeveloperCompetenceData(developerName)['Английский'];
        if (englishLevel == '' || englishLevel == '--') {
            englishLevel =
        }
        reportSheet.getRange(row, 2).setValue(developerName).setVerticalAlignment("middle");
        reportSheet.getRange(row, 3).setValue(englishLevel).setVerticalAlignment("middle");
        // Here you need to calculate trainingAndSales and allocation for each developer
        let trainingHours = developer.projects['Training'] || 0;
        let salesHours = developer.projects['SALES'] || 0;

        reportSheet.getRange(row, 4).setValue(trainingHours).setVerticalAlignment("middle");
        reportSheet.getRange(row, 5).setValue(salesHours).setVerticalAlignment("middle");
        reportSheet.getRange(row, 6).setValue(getAllocationData(developers, developerName)).setVerticalAlignment("middle");
        let resumeLink = getResumeLink(developerName);
        if (resumeLink) {
            reportSheet.getRange(row, 7).setFormula(`=HYPERLINK("${resumeLink}", "Link")`).setVerticalAlignment("middle");
        }
        let competenceText = getDeveloperCompetenceData(developerName)['Инструменты\nБиблиотеки\nСистемы'];
        reportSheet.getRange(row, 8).setValue(getDeveloperCompetenceData(developerName)['Основной стек']).setVerticalAlignment("middle");
        reportSheet.getRange(row, 9).setValue(getDeveloperCompetenceData(developerName)['Дополнительный стек']).setVerticalAlignment("middle").setNote(competenceText);
        reportSheet.getRange(row, 10).setValue(getDeveloperCompetenceData(developerName)['Обучаемость'] + ' ' + getDeveloperCompetenceData(developerName)['Стрессоустойчивость'] + ' ' + getDeveloperCompetenceData(developerName)['Работа в команде'] + ' ' + getDeveloperCompetenceData(developerName)['Работа с клиентом (командой клиента)'] + ' ' + getDeveloperCompetenceData(developerName)['Навыки самопрезентации'] + ' ' + getDeveloperCompetenceData(developerName)['Гибкость мышления']).setVerticalAlignment("middle").setHorizontalAlignment("center");

        let column = 11;
        if (competenceDevelopers.hasOwnProperty(developerName)) {
            for (let competence of competences) {
                let competenceScore = developerData.competences[competence];
                if (competenceScore) {
                    reportSheet.getRange(row, column).setValue(competenceScore.score);
                }
                column++;
            }
        }
        row++;
    }

    // Set the border
    reportSheet.getRange(5, 2, row-5, column-1).setBorder(true, true, true, true, true, true);
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


function getDevelopers(workloadSheet) {
    let developers = [];
    let projects = [];

    let workloadData = workloadSheet.getDataRange().getValues();

    // Retrieve projects from the 5th row
    projects = workloadData[4].slice(1);

    // Iterate through the rows of the workloadData
    for (let i = 5; i < workloadData.length; i++) {
        // Get the developer's name, which is assumed to be in the 4th column
        let developerName = workloadData[i][3];

        // If the developer name is "total", stop the loop
        if (developerName === 'total') {
            break;
        }

        // Create a new developer object
        let developer = {name: developerName, projects: {}};

        // Check if the developer worked on the "Training" or "SALES" projects
        for (let j = 1; j < workloadData[i].length; j++) {
            let hours = workloadData[i][j];
            let projectName = projects[j - 1];

            // If the developer worked on this project, add it to the developer's projects
            if (hours && (projectName === "Training" || projectName === "SALES")) {
                developer.projects[projectName] = hours;
                // Logger.log(projectName + " : " + hours);
            }
        }

        // Only add the developer to the array if they worked on the "Training" or "SALES" projects
        if (Object.keys(developer.projects).length > 0) {
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



function getScrumFilesData(lastWeekMondayDate, lastWeekSundayDate) {
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


// Словарь для перевода английских имен в русские
const nameTranslations = {
    "Aleqsanyan Hovhannes": "Алексанян Ованес",
    "Alexandr Naydenov": "Александр Найденов",
    "Alexeev Alexander": "Алексеев Александр",
    "Alexseev Andrey": "Алексеев Андрей",
    "Alferov Aleksei": "Алферов Алексей",
    "Alputov Anton": "Алпутов Антон",
    "Amandos Manas": "Амандос Манас",
    "Anisimov Ivan": "Анисимов Иван",
    "Antonov Maxim": "Антонов Максим",
    "Arefkin Sergey": "Арефкин Сергей",
    "Baranov Andrey": "Баранов Андрей",
    "Bastov Yegor": "Бастов Егор",
    "Bersenev Nikolay": "Берсенев Николай",
    "Boginich Roman": "Богинич Роман",
    "Borzov Victor": "Борзов Виктор",
    "Brusentsov Dmitry": "Брусенцов Дмитрий",
    "Butyugova Irina": "Бутюгова Ирина",
    "Chikunov Anton": "Чикунов Антон",
    "Chikunov Nikita": "Чикунов Никита",
    "Chizhevsky Dmitrii": "Чижевский Дмитрий",
    "Dombrovsky Nikita": "Домбровский Никита",
    "Dreko Dmitry": "Дреко Дмитрий",
    "Dyatlov Georgy": "Дятлов Георгий",
    "Gasparyan Marine": "Гаспарян Марине",
    "Gilyazov Roman": "Гилязов Роман",
    "Glazachev Vladimir": "Глазачев Владимир",
    "Gulyaev Vitaly": "Гуляев Виталий",
    "Hafizov Abdullo": "Хафизов Абдулло",
    "Harutyunyan Artur": "Арутюнян Артур",
    "Iarovoi Kirill": "Яровой Кирилл",
    "Iordanov Roman": "Иорданов Роман",
    "Kiselev Vadim": "Киселев Вадим",
    "Kocharyan Aram": "Кочарян Арам",
    "Konovalchik Denis": "Коновальчик Денис",
    "Krasilnikov Alexander": "Красильников Александр",
    "Kulemin Alexander": "Кулёмин Александр",
    "Lesnoi Sergei": "Лесной Сергей",
    "Lisitsyn Anton": "Лисицын Антон",
    "Malochkina Natalia": "Малочкина Наталья",
    "Malygin Maxim": "Малыгин Максим",
    "Mukhametshin Denis": "Мухаметшин Денис",
    "Nasibullin Damir": "Насибуллин Дамир",
    "Nikitin Sergey": "Никитин Сергей",
    "Ognev Vladislav": "Огнев Владислав",
    "Olkhovy Sergey": "Ольховый Сергей",
    "Osmanov Servin": "Османов Сервин",
    "Ovcharov Viktor": "Овчаров Виктор",
    "Panfilov Alexander": "Панфилов Александр",
    "Papachristos Alexander": "Папахристос Александр",
    "Petrov Yegor": "Петров Егор",
    "Pobezhimov Alexander": "Побежимов Александр",
    "Ponomarev Mikhail": "Пономарев Михаил",
    "Potapova Ekaterina": "Потапова Екатерина",
    "Rassolenko Anton": "Рассоленко Антон",
    "Ratkin Ivan": "Раткин Иван",
    "Safin Andrey": "Сафин Андрей",
    "Samokhin Daniil": "Самохин Даниил",
    "Shelikhov Dmitriy": "Шелихов Дмитрий",
    "Shilov Dmitrii": "Шилов Дмитрий",
    "Shub Arsenii": "Шуб Арсений",
    "Sirojiddinov Mukhiddin": "Сироджиддинов Мухиддин",
    "Slavgorodsky Ivan": "Славгородский Иван",
    "Slepnev Nikita": "Слепнев Никита",
    "Sofin Vladimir": "Софин Владимир",
    "Sotnikov Gennady": "Сотников Геннадий",
    "Strygin Andrey": "Стрыгин Андрей",
    "Subkhonkulov Furkat": "Субхонкулов Фуркат",
    "Sukhanov Vladimir": "Суханов Владимир",
    "Tolomanenko Vladislav": "Толоманенко Владислав",
    "Trofimenko Anton": "Трофименко Антон",
    "Tulubaev Mikhail": "Тюлюбаев Михаил",
    "Tumakov Andrew": "Тумаков Андрей",
    "Ugarov Rodion": "Угаров Родион",
    "Vlasov Oleg": "Власов Олег",
    "Yakovenko Roman": "Яковенко Роман",
    "Zhuvagin Rodion": "Жувагин Родион"
};


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





