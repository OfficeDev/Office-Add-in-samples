// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

Office.onReady(() => {
    // If needed, Office.js is ready to be called
});

// Your function must be in the global namespace.
function getData(event) {

    // Implement your custom code here. The following code is a simple example.  
    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
