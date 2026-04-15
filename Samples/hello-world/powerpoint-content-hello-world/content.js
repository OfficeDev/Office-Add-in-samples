(function () {
    "use strict";

    // The onReady function must be run each time a new page is loaded.
    Office.onReady(function (info) {
        document.getElementById("get-data-from-selection").addEventListener("click", () => getDataFromSelection());
    });

    // Gets and displays some details about the current slide.
    function getDataFromSelection() {
        if (Office.context.document.getSelectedDataAsync) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        document.getElementById("selected-data").textContent = 'Hello, world! Some slide details are: ' + JSON.stringify(result.value);
                    } else {
                        document.getElementById("selected-data").textContent = 'Error getting slide details.';
                        console.error('Error:', result.error.message);
                    }
                });
        } else {
            document.getElementById("selected-data").textContent = 'Error: Getting slide details isn\'t supported by this host application.';
            console.error('Error:', 'Getting slide details isn\'t supported by this host application.');
        }
    }
})();