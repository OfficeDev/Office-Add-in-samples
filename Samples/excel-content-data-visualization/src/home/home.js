(function () {
    "use strict";

    // The onReady function must be run each time a new page is loaded.
    Office.onReady(function (info) {
        shared.initialize();

        displayDataOrRedirect();
    });

    // Checks if a binding exists, and either displays the visualization
    //        or redirects to the data-binding page.
    async function displayDataOrRedirect() {
        try {
            await Excel.run(async (context) => {
                const binding = context.workbook.bindings.getItemOrNullObject(shared.bindingID);
                await context.sync();

                if (binding.isNullObject) {
                    window.location.href = 'data-binding.html';
                    return;
                }

                // Display data immediately, then register for changes.
                await displayDataForBinding(context, binding);

                binding.onDataChanged.add(async () => {
                    await Excel.run(async (ctx) => {
                        const b = ctx.workbook.bindings.getItem(shared.bindingID);
                        await displayDataForBinding(ctx, b);
                    });
                });
                await context.sync();
            });
        } catch (error) {
            window.location.href = 'data-binding.html';
        }
    }

    // Queries the binding for its data, then delegates to the visualization script.
    async function displayDataForBinding(context, binding) {
        const table = binding.getTable();
        const headerRange = table.getHeaderRowRange();
        const bodyRange = table.getDataBodyRange();
        headerRange.load("values");
        bodyRange.load("values");
        await context.sync();

        // Build a data object compatible with the visualization.display function.
        const headers = [headerRange.values[0]];
        const rows = bodyRange.values;

        const data = { headers: headers, rows: rows };
        visualization.display(document.getElementById('display-data'), data, showError);

        function showError(message) {
            document.getElementById('display-data').innerHTML =
                '<div class="notice">' +
                '    <h3>Error</h3>' + $('<p/>', { text: message })[0].outerHTML +
                '    <a href="data-binding.html">' +
                '        <b>Bind to a different data range?</b>' +
                '    </a>' +
                '</div>';
        }
    }
})();