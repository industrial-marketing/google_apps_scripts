function doGet(request) {
    var name = request.parameter.name;
    var url = request.parameter.url;

    // для тестирования )))
    // var name = 'Алексанян';
    // var url = 'https://docs.google.com/spreadsheets/d/1CkwxtuAdTq_iWNx1d0ugPlzZSdCKiNAq1XlPbEBR0uo/edit#gid=174549316';

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
                // Get the data starting from the next row and current column
                var dataRange = jiraCell.offset(1, 0, sheet.getLastRow() - jiraCell.getRow() + 1, 5);
                var data = dataRange.getValues();

                data = data.filter(function(row) {
                    // Check if the row is empty by checking each column in the row
                    // The every method will return true if all elements in the row are empty or undefined
                    return !row.every(function(cell) {
                        return cell === "" || cell === undefined;
                    });
                });


                var mappedData = data.map(function(row) {
                    // Пропускаем первый элемент массива
                    var newRow = row.slice(1);

                    if (row[4] > 0) {
                        // Заменяем четвертый элемент на пустое значение
                        newRow[0] = row[1];
                        newRow[1] = row[2];
                        newRow[2] = row[3];
                        newRow[3] = '';
                        newRow[4] = row[4];

                        return newRow;
                    }
                });

                // Удаление пустых элементов из результирующего массива
                mappedData = mappedData.filter(function(row) {
                    return row !== undefined;
                });

                // Вывод в консоль
                mappedData.forEach(function(row) {
                    console.log(row);
                });

                data = mappedData;

                if (data.some(item => item === null)) {
                    var errorMessage = "Data is empty or wrong!";
                    Logger.log(errorMessage);
                    return HtmlService.createHtmlOutput(errorMessage);
                }


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

                // Get all sheets in the spreadsheet
                var sheets = externalSpreadsheet.getSheets();

                // Search for an existing sheet with the same name as the month
                var existingSheet = null;
                for (var i = 0; i < sheets.length; i++) {
                    if (sheets[i].getName() === month) {
                        existingSheet = sheets[i];
                        break;
                    }
                }

                var htmlOutput = "";

                if (existingSheet) {
                    // If the sheet already exists, use it
                    newSheet = existingSheet;
                    htmlOutput += "Использован текущий лист " + month + "<br>";
                    // Clear the contents of the existing sheet for columns A, B, C, D, E starting from row 2
                    newSheet.getRange("A2:E").clearContent();
                } else {
                    // If the sheet doesn't exist, create a new sheet
                    var previousSheet = sheets[sheets.length - 1]; // Select the last sheet in the list
                    newSheet = previousSheet.copyTo(externalSpreadsheet);
                    newSheet.setName(month);
                    htmlOutput += "Создан новый лист " + month + "<br>";
                    // Clear the contents of the new sheet for columns A, B, C, D, E starting from row 2
                    newSheet.getRange("A2:E").clearContent();
                }

                try {
                    // Insert the data into the target sheet starting from cell A2
                    var targetRange = newSheet.getRange(2, 1, data.length, 5);
                    targetRange.setValues(data);

                    htmlOutput += "<table>";
                    data.forEach(function(row) {
                        htmlOutput += "<tr>";
                        row.forEach(function(cell) {
                            htmlOutput += "<td>" + cell + "</td>";
                        });
                        htmlOutput += "</tr>";
                    });
                    htmlOutput += "</table>";

                    // Вывод в консоль
                    Logger.log(htmlOutput);

                    return HtmlService.createHtmlOutput(htmlOutput + "Обновлен scrumfile: " + name + " " + url);
                } catch (error) {
                    // Handle the error and display the problematic data in the HTML response
                    var errorMessage = "An error occurred while inserting the data:<br>";
                    errorMessage += "Error message: " + error.message + "<br><br>";
                    errorMessage += "All Data:<br>";

                    htmlOutput += "<table>";
                    data.forEach(function(row) {
                        htmlOutput += "<tr>";
                        row.forEach(function(cell) {
                            htmlOutput += "<td>" + cell + "</td>";
                        });
                        htmlOutput += "</tr>";
                    });
                    htmlOutput += "</table>";

                    errorMessage += htmlOutput;

                    return HtmlService.createHtmlOutput(errorMessage);
                }

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

