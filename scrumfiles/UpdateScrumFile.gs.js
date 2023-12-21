function doGet(request) {
    var name = request.parameter.name;
    var url = request.parameter.url;

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Open the spreadsheet using the provided URL
    var externalSpreadsheet = SpreadsheetApp.openByUrl(url);

    // Get the Scrum report sheet
    var sheet = spreadsheet.getSheetByName("Scrum report");

    // Find the cell in the 5th row that matches the name
    var nameCell = sheet.getRange(5, 1, 1, sheet.getLastColumn()).createTextFinder(name).findNext();

    if (nameCell) {
        // Get the cell below the name cell
        var urlCell = nameCell.offset(1, 0);

        // Verify that the cell below contains a link
        if (urlCell.getValue().startsWith('https://docs.google.com/spreadsheets/d/')) {
            // Verify that the cell under the URL cell contains "Jira"
            var jiraCell = urlCell.offset(1, 0);
            if (jiraCell.getValue() === "Jira") {

                // Extract the month from the last row of data
                var lastRow = data[0];
                var date = lastRow[0];
                var formattedDate = Utilities.formatDate(new Date(date), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");

                // Check for "empty" date
                if (formattedDate === "01/01/1970") {
                    var errorMessage = "Invalid date encountered, defaulting to 01/01/1970";
                    Logger.log(errorMessage);
                    return HtmlService.createHtmlOutput(errorMessage);
                }

                var month = new Date(formattedDate.split('/')[2], formattedDate.split('/')[1] - 1, formattedDate.split('/')[0]).toLocaleString('ru-RU', { month: 'long' }).charAt(0).toUpperCase() + new Date(formattedDate.split('/')[2], formattedDate.split('/')[1] - 1, formattedDate.split('/')[0]).toLocaleString('ru-RU', { month: 'long' }).slice(1);

                // Open the target yearly spreadsheet
                var yearlySpreadsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1zPZJqgrywzb4XkUX_w3W24NPJ5UxdyieEaenOSaWBIs/edit');
                var yearlySheet = yearlySpreadsheet.getSheets()[0];

                // Identify rows that need to be deleted
                var rowsToDelete = [];
                var allData = yearlySheet.getDataRange().getValues();
                allData.forEach(function(row, index) {
                    var rowName = row[0];
                    var rowDate = new Date(row[1]);
                    var rowMonth = rowDate.toLocaleString('ru-RU', { month: 'long' });

                    if (rowName === name && rowMonth === month) {
                        rowsToDelete.push(index + 1);  // add 1 because array indices are 0-based but spreadsheet rows are 1-based
                    }
                });

                var htmlOutput = "";

                // Delete the identified rows, starting from the bottom to avoid offsetting row indices
                if (rowsToDelete.length > 0) {
                    htmlOutput += "Удалены строки: <br>";
                    for (var i = rowsToDelete.length - 1; i >= 0; i--) {
                        yearlySheet.deleteRow(rowsToDelete[i]);
                        htmlOutput += "Строка " + rowsToDelete[i] + "<br>";
                    }
                } else {
                    htmlOutput += "Строки для удаления не найдены.<br>";
                }

                // Append the new data to the yearly spreadsheet
                var data = sheet.getRange("B2:E" + sheet.getLastRow()).getValues();
                var combinedData = data.map(row => [name, ...row]);

                try {
                    yearlySheet.getRange(yearlySheet.getLastRow() + 1, 1, combinedData.length, combinedData[0].length).setValues(combinedData);

                    htmlOutput += "Добавлены строки: <br>";
                    htmlOutput += "<table>";
                    combinedData.forEach(function(row) {
                        htmlOutput += "<tr>";
                        row.forEach(function(cell) {
                            htmlOutput += "<td>" + cell + "</td>";
                        });
                        htmlOutput += "</tr>";
                    });
                    htmlOutput += "</table>";
                } catch (error) {
                    htmlOutput += "Произошла ошибка при добавлении данных: " + error.message;
                }

                // Вывод в консоль
                Logger.log(htmlOutput);

                return HtmlService.createHtmlOutput(htmlOutput);


            } else {
                return HtmlService.createHtmlOutput("The cell below the URL does not contain 'Jira'.");
            }
        } else {
            return HtmlService.createHtmlOutput("The cell below the name does not contain a link.");
        }
    } else {
        return HtmlService.createHtmlOutput("The name '" + name + "' was not found in the Scrum report.");
    }
}
