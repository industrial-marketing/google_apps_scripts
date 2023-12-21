function saveOAuthToken() {
    var token = ScriptApp.getOAuthToken();
    PropertiesService.getScriptProperties().setProperty('OAUTH_TOKEN', token);
}


function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Обновление')
        .addItem('Получить токен', 'saveOAuthToken')
        .addItem('Generate scrum report', 'generateScrumReport')
        .addItem('Обновить "Scrum files for current week"', 'gatherDataInSheetCheckOAuth')
        .addItem('Обновить "Scrum files for last week"', 'gatherDataInSheetLastWeekCheckOAuth')
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


function gatherScrumFilesDataFromFolder(folderId) {
    //if(!folderId) folderId = "1dFHhx4rYzTqnIbw7WzRYsu2nkDD6i_6s"; // папка 2023 - test
    if(!folderId) folderId = "1sRo3anz8iG98SW4EXEHHutgyIZGhUtFq"; // папка 2023
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();
    const files = folder.getFiles();

    // Check if the folder name is a valid year
    // if (folderName.length !== 4 || !folderName.startsWith('20')) {
    //   Logger.log(`Folder name ${folderName} is not a valid year.`);
    //   return;
    // }

    const year = parseInt(folderName);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Scrum files " + folderName + ".10";

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
        // const monthNames = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
        // const monthNames = ["Октябрь"];
        const monthNames = ["", "", "", "", "", "", "", "", "", "Октябрь", "", ""];
        const monthSheetsToProcess = monthNames.filter(monthName => externalFile.getSheetByName(monthName));

        Logger.log(`Processing file ${fileName}.`);
        for (let monthSheetName of monthSheetsToProcess) {
            const monthIndex = monthNames.indexOf(monthSheetName) + 1; // Get the month number
            const externalSheet = externalFile.getSheetByName(monthSheetName);
            const lastRow = externalSheet.getLastRow();
            const monthSheetData = externalSheet.getRange(2, 1, lastRow - 1, 5).getValues();
            const preparedData = []; // To collect data for a month

            Logger.log(`Processing sheet ${monthSheetName} in file ${fileName}.`);
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
                        Logger.log(`Data in row does not belong to the year and month. Skipping data.`);
                    }
                }
                Logger.log('Rows in list ' + preparedData.length);
            });
            // Write all data for the month at once
            if (preparedData.length > 0) {
                Logger.log('TOTAL Rows in list ' + preparedData.length);
                const startRow = scrumFilesSheet.getLastRow() + 1; // Get the starting row
                scrumFilesSheet.getRange(startRow, 1, preparedData.length, headers.length).setValues(preparedData);
            }
        }
    }

    // Add date and time of data gathering
    const currentTime = new Date().toLocaleString("en-GB", {timeZone: "Asia/Tbilisi"});
    scrumFilesSheet.getRange("A1").setValue(`Data gathered at ${currentTime} (Tbilisi, Georgia Timezone)`);

    Logger.log(`Data gathering finished for ${sheetName}.`);
}


function exportScrumFilesDataToFolder(folderId) {
    if (!folderId) folderId = "1dFHhx4rYzTqnIbw7WzRYsu2nkDD6i_6s";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Scrum files 2023");
    const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, 5).getValues();
    const templateId = "1p5eq0Z_NNKdnKSrmwOAlCyzeNHrIG3kI75f4cHNbM2w";
    const folder = DriveApp.getFolderById(folderId);

    const groupedByDeveloper = groupByDeveloper(data); // Group by Developer

    // First create all necessary files
    for (let developer in groupedByDeveloper) {
        Logger.log(`Checking file for developer ${developer}.`);
        // Check if the developer's file exists, and create it if not
        getFileForDeveloperInFolder(developer, folder, templateId);
    }

    // Then process them, starting with the least recently modified one
    const files = folder.getFiles();
    const filesArray = [];
    while (files.hasNext()) {
        const file = files.next();
        filesArray.push(file);
    }
    filesArray.sort((a, b) => a.getLastUpdated() - b.getLastUpdated());

    for (const file of filesArray) {
        const developer = file.getName().split(".")[0]; // Get developer name from filename by removing the extension

        if (!groupedByDeveloper.hasOwnProperty(developer)) {
            continue;
        }

        Logger.log(`Processing developer ${developer}.`);

        const developerData = groupedByDeveloper[developer];
        const monthNames = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
        const groupedByMonth = groupByMonth(developerData); // only for determining the sheet names

        for (let monthIndex = 0; monthIndex < 12; monthIndex++) {
            if (!groupedByMonth.hasOwnProperty(monthIndex)) {
                continue;
            }
            let month = monthNames[monthIndex];
            Logger.log(`Processing month ${month} for developer ${developer}.`);

            const fileSpreadsheet = SpreadsheetApp.openById(file.getId());
            let monthSheet = fileSpreadsheet.getSheetByName(month);
            if (!monthSheet) {
                // Copy the first sheet (January) if the sheet for the month doesn't exist
                let januarySheet = fileSpreadsheet.getSheets()[0];
                monthSheet = januarySheet.copyTo(fileSpreadsheet);
                monthSheet.setName(month);
                Logger.log(`Sheet for month ${month} created.`);
            }

            // Filter the original data for the specific month and developer, without grouping
            const monthData = developerData
                .filter(row => {
                    // Get the month and year from the date string
                    const date = new Date(row[1]);
                    // Split the date string and rearrange it to MM/dd/yyyy
                    row[1] = Utilities.formatDate(date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
                    const [day, month, year] = row[1].split("/");
                    row[1] = `${month}/${day}/${year}`;
                    // Check if the month of the date matches the current month
                    return date.getMonth() === monthIndex;
                })
                .map(row => [row[1], row[2], row[3], "", row[4]])
                .sort((a, b) => new Date(a[0]) - new Date(b[0])); // Sort by date

            // Get current data from the sheet
            let currentData = monthSheet.getRange(2, 1, monthSheet.getLastRow() - 1, 5).getValues();
            // Filter the original data for the specific month and developer, without grouping
            currentData = currentData
                .filter(row => {
                    // Get the month and year from the date string
                    const date = new Date(row[0]);
                    // Split the date string and rearrange it to MM/dd/yyyy
                    row[0] = Utilities.formatDate(date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
                    const [day, month, year] = row[0].split("/");
                    row[0] = `${month}/${day}/${year}`;
                    // Check if the month of the date matches the current month
                    return date.getMonth() === monthIndex;
                })
                .sort((a, b) => new Date(a[0]) - new Date(b[0])); // Sort by date

            // Check if the data has changed
            const dataHasChanged = JSON.stringify(monthData) !== JSON.stringify(currentData);

            if (dataHasChanged) {
                // Clear old data and insert new data
                monthSheet.getRange("A2:E").clearContent();
                if (monthData.length > 0) {
                    monthSheet.getRange(2, 1, monthData.length, 5).setValues(monthData);
                    Logger.log(`Data for month ${month} inserted/updated.`);
                } else {
                    Logger.log(`No data for month ${month}.`);
                }
            } else {
                Logger.log(`Data for month ${month} has not changed.`);
            }
        }
    }

    Logger.log(`Finished exporting data to folder ${folderId}.`);
}


function getFileForDeveloperInFolder(developer, folder, templateId) {
    const fileIterator = folder.searchFiles(`title = '${developer}'`);
    if (fileIterator.hasNext()) {
        return SpreadsheetApp.openById(fileIterator.next().getId());
    } else {
        // Copy template if file doesn't exist
        const template = DriveApp.getFileById(templateId);
        const newFile = template.makeCopy(developer, folder);
        return SpreadsheetApp.openById(newFile.getId());
    }
}

function groupByMonth(data) {
    return data.reduce((groups, row) => {
        const date = new Date(row[1]);
        const monthIndex = date.getMonth(); // month is 0-indexed
        if (!groups[monthIndex]) {
            groups[monthIndex] = [];
        }
        groups[monthIndex].push(row);
        return groups;
    }, {});
}

function groupByDeveloper(data) {
    return data.reduce((groups, row) => {
        const developer = row[0];
        if (!groups[developer]) {
            groups[developer] = [];
        }
        groups[developer].push(row);
        return groups;
    }, {});
}






function generateScrumReport() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    var settingsSheet = spreadsheet.getSheetByName('Scrum report settings');
    var scrumFilesSheet = spreadsheet.getSheetByName('Scrum files');
    var reportSheet = spreadsheet.getSheetByName('Scrum report');
    var studentsSheet = spreadsheet.getSheetByName('Developers');
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
                // Logger.log(date);
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
            //Logger.log(date);
        }

        // Logger.log(date);

        if(!date) continue;

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
        var hyperlinkUrl = "https://script.google.com/a/macros/sharp-dev.net/s/AKfycbwqiUBVMxL0VXe8weKfn_RLwnZX9A1M8Sfm6_CsFS8dOjTd6fjDgB2m99xKGcZVPL5TUw/exec?name=" + encodedName + "&url=" + encodedFileLink;
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


