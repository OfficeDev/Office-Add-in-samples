(function () {
    "use strict";

    // The onReady function must be run each time a new page is loaded.
    Office.onReady(function (info) {
        document.getElementById("get-data-from-selection").addEventListener("click", () => getDataFromSelection());
    });

    // Reads data from current document selection and displays it.
    function getDataFromSelection() {
        if (Office.context.document.getSelectedDataAsync) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        document.getElementById("selected-data").textContent = 'Hello, world! The selected text is: ' + result.value;
                    } else {
                        document.getElementById("selected-data").textContent = 'Error getting selected text.';
                        console.error('Error:', result.error.message);
                    }
                });
        } else {
            document.getElementById("selected-data").textContent = 'Error: Reading selection data isn\'t supported by this host application.';
            console.error('Error:', 'Reading selection data isn\'t supported by this host application.');
        }
    }
})();