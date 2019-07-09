'use strict';

(function () {

    // Office is ready
    Office.onReady(function () {
        // The document is ready
        $(document).ready(function () {

            // Assign HTML buttons to functions
            $('#create-table').click(loadTable);
        });
    });

    function loadTable() {
        if (localStorage.DraftPlayerData) {
            var dataObject = JSON.parse(localStorage.DraftPlayerData);
            createTable(dataObject);
        }
        else {
            $.ajax({
                dataType: "json",
                url: "sampleData.js",
                success: function (result, status, xhr) {
                    localStorage.DraftPlayerData = JSON.stringify(result);
                    createTable(result);
                },
                error: function (xhr, status, error) {
                    console.log("Player data failed to load with error: " + error);
                }
            });
        }
    }

    function createTable(playerData) {
        Excel.run(function (context) {
            var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            var playerTable = currentWorksheet.tables.getItemOrNullObject('PlayerTable');

            return context.sync().then(function () {
                playerTable.delete();

                playerTable = currentWorksheet.tables.add("A1:D1", true);
                playerTable.name = "PlayerTable";

                playerTable.getHeaderRowRange().values =
                    [["Name", "PPG", "Rebounds", "APG"]];

                playerTable.rows.add(null, playerData);

                playerTable.getRange().format.autofitColumns();
                playerTable.getRange().format.autofitRows();

                return context.sync();
            });
        }).catch(errorHandleFunction);
    }

    function errorHandleFunction(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

})();