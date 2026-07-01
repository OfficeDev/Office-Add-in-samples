(function () {
    'use strict';

    // The onReady function must be defined for each page in your add-in.
    Office.onReady(function (info) {
        shared.initialize();

        document.getElementById("bind-to-existing-data").addEventListener("click", () => bindToExistingData());
        document.getElementById("insert-sample-data").addEventListener("click", () => insertSampleData());
        document.getElementById("insert-data-available").style.display = "block";
        document.getElementById("insert-data-unavailable").style.display = "none";
    });

    // Binds the visualization to existing data.
    // Note: addFromPromptAsync has no application-specific API equivalent
    // because it provides a built-in range picker dialog.
    function bindToExistingData() {
        Office.context.document.bindings.addFromPromptAsync(
            Office.BindingType.Table,
            { id: shared.bindingID, sampleData: visualization.generateSampleData() },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    window.location.href = 'home.html';
                } else {
                    shared.showNotification(result.error.name, result.error.message);
                }
            }
        );
    }

    // Inserts sample data into the current selection and creates a binding.
    async function insertSampleData() {
        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                const sampleData = visualization.generateSampleData();

                // Build a new range to fit the sample data and insert it into the worksheet.
                const values = sampleData.headers.concat(sampleData.rows);
                const targetRange = range.getResizedRange(
                    values.length - 1,
                    values[0].length - 1
                );
                targetRange.values = values;

                // Create a table from the inserted data and bind to it.
                const table = context.workbook.tables.add(targetRange, true /* hasHeaders */);
                table.name = shared.bindingID;

                await context.sync();
                window.location.href = 'home.html';
            });
        } catch (error) {
            shared.showNotification('Error', error.message);
        }
    }
})();
