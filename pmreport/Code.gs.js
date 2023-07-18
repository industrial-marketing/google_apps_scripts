function saveOAuthToken() {
    var token = ScriptApp.getOAuthToken();
    PropertiesService.getScriptProperties().setProperty('OAUTH_TOKEN', token);
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
        .addItem('Save OAuth Token', 'saveOAuthToken')
        .addItem('Generate last week report for current PM', 'generateLastWeekReportForCurrentPMCheckOAuth')
        .addItem('Gather scrum files data for last week in "Scrum files for last week"', 'gatherDataInCurrentSheetCheckOAuth')
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

function generateLastWeekReportForCurrentPMCheckOAuth() {
    // Проверка наличия OAuth-токена
    var hasToken = checkOAuthToken();

    if (hasToken) {
        // Токен присутствует, выполняем функцию updateWeekPlan()
        generateLastWeekReportForCurrentPM();
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

function generateLastWeekReportForCurrentPM() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const pmLastName = activeSheet.getName();

    const pmListSheet = ss.getSheetByName("PM list");
    const pmList = pmListSheet.getRange("A2:B" + pmListSheet.getLastRow()).getValues();
    const pm = pmList.find(row => row[0] === pmLastName && row[1]);

    if (!pm) {
        SpreadsheetApp.getUi().alert('Please check the current sheet. It should be a project manager\'s report sheet.');
        return;
    }

    const pmInitials = pm[1];

    const reportSheetId = "1N65NUtqBA855C6K8swmeFQ9HbvIZU4fq4EnhYzvNV7Q";
    const reportSpreadsheet = SpreadsheetApp.openById(reportSheetId);
    const lastWeekMondayDate = new Date();
    lastWeekMondayDate.setDate(lastWeekMondayDate.getDate() - lastWeekMondayDate.getDay() - 6);
    const lastWeekSundayDate = new Date();
    lastWeekSundayDate.setDate(lastWeekSundayDate.getDate() - lastWeekSundayDate.getDay());
    const lastWeekMondayString = Utilities.formatDate(lastWeekMondayDate, ss.getSpreadsheetTimeZone(), 'd MMM').toLowerCase();
    const lastWeekSundayString = Utilities.formatDate(lastWeekSundayDate, ss.getSpreadsheetTimeZone(), 'd MMM').toLowerCase();

    let reportSheetName = lastWeekMondayDate.getMonth() === lastWeekSundayDate.getMonth() ?
        `${lastWeekMondayString.split(" ")[0]}-${lastWeekSundayString.split(" ")[0]} ${lastWeekSundayString.split(" ")[1]}` :
        `${lastWeekMondayString}-${lastWeekSundayString.split(" ")[0]}`;

    const reportSheet = reportSpreadsheet.getSheetByName(reportSheetName);
    if (!reportSheet) {
        SpreadsheetApp.getUi().alert(`Cannot find sheet "${reportSheetName}" in the report spreadsheet.`);
        return;
    }

    const scrumSheet = ss.getSheetByName('Scrum files for last week');
    let scrumData;
    try {
        const range = scrumSheet.getRange('A3:E' + scrumSheet.getLastRow());
        scrumData = range.getValues();
    } catch (error) {
        SpreadsheetApp.getUi().alert("Error retrieving data from the scrum sheet: " + error);
        return;
    }

    const filteredScrumData = scrumData.filter(row => {
        const date = row[1]; // Предполагается, что дата находится во второй колонке (B)
        return date >= lastWeekMondayDate && date <= lastWeekSundayDate;
    });

    if (filteredScrumData.length === 0) {
        SpreadsheetApp.getUi().alert("No matching data found in 'Scrum files for last week' for dates " + lastWeekMondayString + "-" + lastWeekSundayString);
        return;
    }

    let lastReportRow = activeSheet.getLastRow();
    let totalHours = 0;

    for (let i = lastReportRow; i > 0; i--) {
        const weekCell = activeSheet.getRange(`B${i}`);
        const weekValue = weekCell.getValue();
        if (weekValue === `${Utilities.formatDate(lastWeekMondayDate, ss.getSpreadsheetTimeZone(), 'd.MM.yyyy')} - ${Utilities.formatDate(lastWeekSundayDate, ss.getSpreadsheetTimeZone(), 'd.MM.yyyy')}`) {
            const toDeleteRow = weekCell.getRow();
            const deletedRowsCount = lastReportRow - toDeleteRow + 1;
            activeSheet.deleteRows(toDeleteRow, deletedRowsCount);
            lastReportRow = activeSheet.getLastRow();
            break;
        }
    }

    activeSheet.getRange(`A${lastReportRow + 6}`).setValue('Week');
    activeSheet.getRange(`B${lastReportRow + 6}`).setValue(`${Utilities.formatDate(lastWeekMondayDate, ss.getSpreadsheetTimeZone(), 'd.MM.yyyy')} - ${Utilities.formatDate(lastWeekSundayDate, ss.getSpreadsheetTimeZone(), 'd.MM.yyyy')}`);
    activeSheet.getRange(`A${lastReportRow + 8}`).setValue(pmInitials);

    const headersRange = reportSheet.getRange("1:1");
    const headers = headersRange.getValues()[0];
    const pmColumn = headers.indexOf('Project manager') + 1;
    const emptyColumn = headers.slice(5).indexOf('') + 6;

    if (pmColumn === 0 || emptyColumn === 0) {
        SpreadsheetApp.getUi().alert(`Error reading data from the report spreadsheet.`);
        return;
    }

    const reportData = reportSheet.getRange(1, pmColumn, reportSheet.getLastRow(), emptyColumn - pmColumn).getValues();
    const projectNames = reportSheet.getRange(5, pmColumn, 1, emptyColumn - pmColumn).getValues()[0];

    let scrumDataGroupedByDeveloperAndProject = {};
    let scrumDataGroupedByDeveloper = {};
    let scrumDataGroupedByProject = {};
    scrumData.forEach(row => {
        if (row[2] == "HR") row[3] = "HR";
        if (row[2] == "PRESALE") row[3] = "SALES";
        if (row[2] == "Administrative") row[3] = "Administrative";
        if (row[2] == "Testing") row[3] = "Testing";
        if (row[2] == "DevOps") row[3] = "DevOps";
        const [developer, date, type, project, hours] = row;
        if (!scrumDataGroupedByDeveloperAndProject[developer]) {
            scrumDataGroupedByDeveloperAndProject[developer] = {};
            scrumDataGroupedByDeveloper[developer] = {};
        }
        if (!scrumDataGroupedByDeveloperAndProject[developer][project]) {
            scrumDataGroupedByDeveloperAndProject[developer][project] = {
                totalHours: 0,
                nonBillableHours: 0
            };
        }
        if (!scrumDataGroupedByDeveloper[developer][project]) {
            scrumDataGroupedByDeveloper[developer][project] = 0;
        }
        if (!scrumDataGroupedByProject[project]) {
            scrumDataGroupedByProject[project] = {
                totalHours: 0,
                nonBillableHours: 0
            };
        }
        scrumDataGroupedByDeveloperAndProject[developer][project].totalHours += Math.round(hours * 100) / 100;
        scrumDataGroupedByDeveloper[developer][project] += Math.round(hours * 100) / 100;
        scrumDataGroupedByProject[project].totalHours += Math.round(hours * 100) / 100;
        if (type === 'DEVfree' || type === 'HR' || type === 'PRESALE' || type === 'Administrative' || type === 'Testing' || type === 'DevOps' || type === 'Learning') {
            scrumDataGroupedByDeveloperAndProject[developer][project].nonBillableHours += Math.round(hours * 100) / 100;
            scrumDataGroupedByProject[project].nonBillableHours += Math.round(hours * 100) / 100;
        }
    });

    const headersCells = ["B", "C", "D", "E", "F", "G", "H"];
    const headersNames = ['Project', 'Planned total/ non-billable', 'Fact Total', 'Fact non-billable', 'Fact/Plan difference', 'Actual Allocation', 'PM Comments'];

    headersNames.forEach((header, index) => {
        const cell = activeSheet.getRange(`${headersCells[index]}${lastReportRow + 8}`);
        cell.setValue(header).setFontWeight('bold').setBackground("#cccccc");
    });

    let currentRow = lastReportRow + 9;
    let totalFactTotal = 0;
    let totalFactNonBillable = 0;
    let totalFactPlanDifference = 0;

    for (let i = 0; i < reportData[0].length; i++) {
        if (reportData[0][i] === pmInitials) {
            let projectHours = 0;
            let projectNonBillable = 0;
            let projectDevelopers = [];
            for (let j = 5; j < reportData.length; j++) {
                if (reportData[j][i] > 0) {
                    if (reportData[j][0] === "total") {
                        projectHours += Math.round(reportData[j][i] * 100) / 100;
                        totalHours += Math.round(reportData[j][i] * 100) / 100;
                    } else if (reportData[j][0] === "Бесплатно:") {
                        projectNonBillable += Math.round(reportData[j][i] * 100) / 100;
                        break;
                    } else if (reportData[j][0] === "paid" || reportData[j][0] === "scrum file total 4 days appr to 5") {
                        break;
                    } else {
                        projectDevelopers.push({name: reportData[j][0], hours: Math.round(reportData[j][i] * 100) / 100});
                    }
                }
            }

            if (projectHours > 0) {
                activeSheet.getRange(`B${currentRow}`).setValue(projectNames[i]).setBackground("#d9ead3");
                activeSheet.getRange(`C${currentRow}`).setValue(`${projectHours}/${projectNonBillable}`).setBackground("#d9ead3");
                if (scrumDataGroupedByProject[projectNames[i]]) {
                    activeSheet.getRange(`D${currentRow}`).setValue(scrumDataGroupedByProject[projectNames[i]].totalHours).setBackground("#d9ead3");
                    activeSheet.getRange(`E${currentRow}`).setValue(scrumDataGroupedByProject[projectNames[i]].nonBillableHours).setBackground("#d9ead3");
                    activeSheet.getRange(`F${currentRow}`).setValue(projectHours - scrumDataGroupedByProject[projectNames[i]].totalHours).setBackground("#d9ead3");
                }
                activeSheet.getRange(`G${currentRow}`).setValue('').setBackground("#d9ead3");
                activeSheet.getRange(`H${currentRow}`).setValue('').setBackground("#d9ead3");
                currentRow++;

                let projectName = projectNames[i];

                let scrumDevelopers = Object.keys(scrumDataGroupedByDeveloperAndProject).filter(dev => Object.keys(scrumDataGroupedByDeveloperAndProject[dev]).includes(projectName));

                let allScrumDevelopers = Object.keys(scrumDataGroupedByDeveloperAndProject).filter(dev => Object.keys(scrumDataGroupedByDeveloperAndProject[dev]));

                let projectDevelopersRenamed = []; // projectDevelopers.map(dev => dev.name.split(' ')[0]); // Get short names from the workload report

                projectDevelopers.forEach(workloadDev => {
                    allScrumDevelopers.forEach(scrumDev => {
                        if (workloadDev.name.startsWith(scrumDev)) {
                            projectDevelopersRenamed.push(scrumDev);
                        }
                    });
                });

                let allDevelopers = [...new Set([...projectDevelopersRenamed, ...scrumDevelopers])]; // Now we are comparing short names

                for (let developerShortName of allDevelopers) {
                    let devHours = projectDevelopers.find(dev => dev.name.startsWith(developerShortName))?.hours || 0;

                    let rowFontColor = '#000000';
                    if (developerShortName === pmLastName) {
                        rowFontColor = '#1155cc';
                    }

                    activeSheet.getRange(`B${currentRow}`).setValue(developerShortName).setFontColor(rowFontColor);
                    activeSheet.getRange(`C${currentRow}`).setValue(devHours).setFontColor(rowFontColor);

                    // Adding a check before trying to access properties
                    if (scrumDataGroupedByDeveloper[developerShortName]) {
                        let allocationList = [];
                        let totalDeveloperHours = 0; // Initialize total developer hours here
                        for (let project in scrumDataGroupedByDeveloper[developerShortName]) {
                            let roundedHours = Math.round(scrumDataGroupedByDeveloper[developerShortName][project] * 100) / 100;
                            totalDeveloperHours += roundedHours; // Increment total hours with each project
                            allocationList.push(project + ' (' + roundedHours + ')');
                        }
                        // Append total hours to the end of the list
                        // allocationList.push('TOTAL (' + Math.round(totalDeveloperHours * 100) / 100 + ')');
                        allocationList = Math.round(totalDeveloperHours * 100) / 100 + ': ' + allocationList.join(', ');
                        activeSheet.getRange(`G${currentRow}`).setValue(allocationList).setFontColor(rowFontColor);
                    }


                    // Adding a check before trying to access properties
                    if (scrumDataGroupedByDeveloperAndProject[developerShortName] && scrumDataGroupedByDeveloperAndProject[developerShortName][projectName]) {
                        let totalHours = scrumDataGroupedByDeveloperAndProject[developerShortName][projectName].totalHours;
                        let nonBillableHours = scrumDataGroupedByDeveloperAndProject[developerShortName][projectName].nonBillableHours;
                        let factPlanDifference = devHours - totalHours;

                        activeSheet.getRange(`D${currentRow}`).setValue(totalHours).setFontColor(rowFontColor);
                        activeSheet.getRange(`E${currentRow}`).setValue(nonBillableHours).setFontColor(rowFontColor);
                        activeSheet.getRange(`F${currentRow}`).setValue(factPlanDifference).setFontColor(rowFontColor);

                        totalFactTotal += totalHours;
                        totalFactNonBillable += nonBillableHours;
                        totalFactPlanDifference += factPlanDifference;

                        if (factPlanDifference > 0) {
                            activeSheet.getRange(`F${currentRow}`).setFontColor("#ff0000");
                        }
                    } else {
                        let totalHours = 0;
                        let nonBillableHours = 0;
                        let factPlanDifference = devHours - totalHours;
                        activeSheet.getRange(`D${currentRow}`).setValue(totalHours).setFontColor(rowFontColor);
                        activeSheet.getRange(`E${currentRow}`).setValue(nonBillableHours).setFontColor(rowFontColor);
                        activeSheet.getRange(`F${currentRow}`).setValue(factPlanDifference).setFontColor(rowFontColor);
                        if (factPlanDifference > 0) {
                            activeSheet.getRange(`F${currentRow}`).setFontColor("#ff0000");
                        }
                    }
                    currentRow++;
                }
            }
        }
    }

    activeSheet.getRange(`B${currentRow}`).setValue('Total').setFontWeight('bold');
    activeSheet.getRange(`C${currentRow}`).setValue(totalHours).setFontWeight('bold');
    activeSheet.getRange(`D${currentRow}`).setValue(totalFactTotal).setFontWeight('bold');
    activeSheet.getRange(`E${currentRow}`).setValue(totalFactNonBillable).setFontWeight('bold');
    activeSheet.getRange(`F${currentRow}`).setValue(totalFactPlanDifference).setFontWeight('bold');
    activeSheet.getRange(`B${lastReportRow + 9}:H${currentRow}`).setBorder(true, true, true, true, true, true);
}



function getScrumFilesData(lastWeekMondayDate, lastWeekSundayDate) {
    const scrumFilesFolderId = "1AnMMx9rnnQE7r2KoodgYZP1eukycY0lJ";
    const folder = DriveApp.getFolderById(scrumFilesFolderId);
    const monthNames = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
    const startYearMonthName = `${lastWeekMondayDate.getFullYear()} ${(lastWeekMondayDate.getMonth() + 1).toString().padStart(2, '0')} ${monthNames[lastWeekMondayDate.getMonth()]}`;
    const endYearMonthName = `${lastWeekSundayDate.getFullYear()} ${(lastWeekSundayDate.getMonth() + 1).toString().padStart(2, '0')} ${monthNames[lastWeekSundayDate.getMonth()]}`;

    Logger.log(`Looking for files: ${startYearMonthName} and ${endYearMonthName}`);

    const filesIterator = folder.getFiles();
    let scrumFiles = [];
    while (filesIterator.hasNext()) {
        const file = filesIterator.next();
        Logger.log(`Checking file: ${file.getName()}`);
        if (file.getName() === startYearMonthName || file.getName() === endYearMonthName) {
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
        const sheet = ss.getSheetByName('СкрамФайлы');
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
            const externalSheet = externalFile.getSheetByName('Июль');
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

function getLastSunday() {
    const today = new Date();
    const lastSunday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - today.getDay());
    return lastSunday;
}

function getLastMonday() {
    const today = new Date();
    const lastMonday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - today.getDay() - 6);
    return lastMonday;
}

function gatherDataInCurrentSheet() {
    const lastWeekMondayDate = getLastMonday();
    const lastWeekSundayDate = getLastSunday();

    const data = getScrumFilesData(lastWeekMondayDate, lastWeekSundayDate);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let dataSheet = ss.getSheetByName("Scrum files for last week");

    // Создаем новый лист, если он не существует
    if (!dataSheet) {
        dataSheet = ss.insertSheet("Scrum files for last week");
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

function testGetScrumFilesData() {
    const lastWeekMondayDate = new Date(2023, 6, 10); // Месяцы в JavaScript начинаются с 0, поэтому июль будет 6
    const lastWeekSundayDate = new Date(2023, 6, 16);

    const data = getScrumFilesData(lastWeekMondayDate, lastWeekSundayDate);

    // Запись результатов в лог для проверки
    Logger.log(data);
}

