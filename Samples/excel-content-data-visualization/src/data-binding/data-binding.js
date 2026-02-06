(function () {
    'use strict';

    // The onReady function must be defined for each page in your add-in.
    Office.onReady(function (info) {
        shared.initialize();

        document.getElementById("bind-to-existing-data").addEventListener("click", () => bindToExistingData());

        if (dataInsertionSupported()) {
            document.getElementById("insert-sample-data").addEventListener("click", () => insertSampleData());
            document.getElementById("insert-data-available").style.display = "block";
            document.getElementById("insert-data-unavailable").style.display = "none";
        } else {
            document.getElementById("insert-data-available").style.display = "none";
            document.getElementById("insert-data-unavailable").style.display = "block";
        }
    });

    // Binds the visualization to existing data.
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

    // Checks whether the current application supports setting selected data.
    function dataInsertionSupported() {
        return Office.context.document.setSelectedDataAsync &&
            (Office.context.document.bindings &&
                Office.context.document.bindings.addFromSelectionAsync);
    }

    // Inserts sample data into the current selection (if supported).
    function insertSampleData() {
        Office.context.document.setSelectedDataAsync(visualization.generateSampleData(),
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    Office.context.document.bindings.addFromSelectionAsync(
                        Office.BindingType.Table, { id: shared.bindingID },
                        function (result) {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                window.location.href = 'home.html';
                            } else {
                                shared.showNotification(result.error.name, result.error.message);
                            }
                        }
                    );
                } else {
                    shared.showNotification(result.error.name, result.error.message);
                }
            }
        );
    }
})();
