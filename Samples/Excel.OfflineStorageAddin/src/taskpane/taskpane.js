// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
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
    
    // Loads data for a table from the server if connected
    // or local storage if disconnected and had been loaded before
    function loadTable() {
        $.ajax({
            dataType: "json",
            url: "sampleData.json",
            success: function (result, status, xhr) {
                // Stores the JSON retrieved from the AJAX call as a string in
                // local storage under the key "PlayerData"
                localStorage.PlayerData = JSON.stringify(result);

                // Sends the new data to the table
                createTable(result);
            },
            error: function (xhr, status, error) {
                // If the connections fails, checks if "PlayerData" was previously stored in local storage
                if (localStorage.PlayerData) {
                    // Retrieves the string saved earlied under the key "PlayerData"
                    // and parses it into an object
                    let dataObject = JSON.parse(localStorage.PlayerData);

                    // Sends the saved data to the table
                    createTable(dataObject);
                }
                else {
                    console.log("Player data failed to load with error: " + error);
                }
            }
        });
    }
    
    // Creates and populates table of stats from sampleData.json
    function createTable(playerData) {
        Excel.run(function (context) {
            let currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
            let playerTable = currentWorksheet.tables.getItemOrNullObject('PlayerTable');

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
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
                    
})();
