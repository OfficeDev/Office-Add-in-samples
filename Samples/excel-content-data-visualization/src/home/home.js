(function () {
    "use strict";

    // The onReady function must be run each time a new page is loaded.
    Office.onReady(function (info) {
        shared.initialize();

        displayDataOrRedirect();
    });

    // Checks if a binding exists, and either displays the visualization
    //        or redirects to the data-binding page.
    function displayDataOrRedirect() {
        Office.context.document.bindings.getByIdAsync(
            shared.bindingID,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const binding = result.value;
                    let handler = function () { displayDataForBinding(binding); };
                    binding.addHandlerAsync(
                        Office.EventType.BindingDataChanged,
                        handler,
                        handler
                    );
                } else {
                    window.location.href = 'data-binding.html';
                }
            });
    }

    // Queries the binding for its data, then delegates to the visualization script.
    function displayDataForBinding(binding) {
        binding.getDataAsync(
            {
                coercionType: Office.CoercionType.Table,
                valueFormat: Office.ValueFormat.Unformatted,
                filterType: Office.FilterType.OnlyVisible
            },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    visualization.display(document.getElementById('display-data'), result.value, showError);
                } else {
                    showError('Could not read data.');
                }
            }
        );

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