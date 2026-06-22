(function () {
    "use strict";

    // The onReady function must be run each time a new page is loaded.
    Office.onReady(function (info) {
        document.getElementById("get-data-from-selection").addEventListener("click", () => getDataFromSelection());
    });

    // Reads data from current document selection and displays it.
    async function getDataFromSelection() {
        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.load("text");
                await context.sync();

                // Join multi-cell selections into a single string.
                const text = range.text.map((row) => row.join(", ")).join("; ");
                document.getElementById("selected-data").textContent = 'Hello, world! The selected text is: ' + text;
            });
        } catch (error) {
            document.getElementById("selected-data").textContent = 'Error getting selected text.';
            console.error('Error:', error.message);
        }
    }
})();