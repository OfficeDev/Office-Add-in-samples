// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/// <reference path="../App.js" />

// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons

(function () {

    Office.initialize = function (reason) {
        // Use this to check the APIs supported by Word. 
        if (Office.context.requirements.isSetSupported("WordApi", "1.1")) {
            // Wire up the click events for the two buttons in the WD_OpenXML_js.html page
            // so that they use the new Word API.
            document.getElementById('getOOXMLData').onclick = getOOXML_newAPI;
            document.getElementById('setOOXMLData').onclick = setOOXML_newAPI;
            console.log('This code is using Word 2016 or later.');
        } else {
            // Wire up the click events of the two buttons in the WD_OpenXML_js.html page
            // so that they use the original Office JS API.
            document.getElementById('getOOXMLData').onclick = getOOXML;
            document.getElementById('setOOXMLData').onclick = setOOXML;
            console.log('This code is using Word 2013.');
        }
    };

    // Variable to hold Office Open XML.
    let currentOOXML = "";

    // Gets the OOXML contents of the Word document body and
    // puts the OOXML into a textarea in the add-in.
    function getOOXML_newAPI() {

        // Get a reference to the Div where we will write the status of our operation
        const report = document.getElementById("status");

        // Remove all nodes from the status Div so we have a clean space to write to
        while (report.hasChildNodes()) {
            report.removeChild(report.lastChild);
        }

        // Get a reference to the text area that will hold the OOXML we get from the document.
        const textArea = document.getElementById("dataOOXML");

        // Run a batch operation against the Word Javascript object model.
        Word.run(function (context) {

            // Create a proxy object for the document body.
            const body = context.document.body;

            // Queue a commmand to get the OOXML contents of the body.
            const bodyOOXML = body.getOoxml();

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {

                // Stashing the OOXML in case
                currentOOXML = bodyOOXML.value;

                // Update the status message.
                setTimeout(function () {
                    textArea.value = currentOOXML;
                    report.innerText = "The getOOXML function succeeded!";
                }, 400);

                // Clear the success message after a 2 second delay
                setTimeout(function () {
                    report.innerText = "";
                }, 2000);

            });
        })
            .catch(function (error) {

                // Clear the OOXML, show the error info
                currentOOXML = "";
                report.innerText = error.message;

                console.log("Error: " + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }

    // Gets the OOXML contents in a textarea from the add-in and puts the contents of 
    // into the Word document body. 
    function setOOXML_newAPI() {
        // Get a reference to the Div where we will write the outcome of our operation.
        const report = document.getElementById("status");

        // Sets the currentOOXML variable to the current contents of the task pane text area.
        const textArea = document.getElementById("dataOOXML");
        currentOOXML = textArea.value;

        // Remove all nodes from the status Div so we have a clean space to write to.
        while (report.hasChildNodes()) {
            report.removeChild(report.lastChild);
        }

        // Check whether we have OOXML in the variable.
        if (currentOOXML != "") {

            // Run a batch operation against the Word object model.
            Word.run(function (context) {

                // Create a proxy object for the document body.
                const body = context.document.body;

                // Queue a commmand to insert OOXML in to the beginning of the body.
                body.insertOoxml(currentOOXML, Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync().then(function () {

                    // Tell the user we succeeded and then clear the message after a 2-second delay.
                    report.innerText = "The setOOXML function succeeded!";
                    setTimeout(function () {
                        report.innerText = "";
                    }, 2000);
                });
            })
                .catch(function (error) {

                    // Clear the text area just so we don't give you the impression that there's
                    // valid OOXML waiting to be inserted... 
                    textArea.value = "";
                    // Let the user see the error.
                    report.innerText = error.message;

                    console.log('Error: ' + JSON.stringify(error));
                    if (error instanceof OfficeExtension.Error) {
                        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                    }
                });

        } else {
            report.innerText = 'Add some OOXML data before trying to set the contents.';
        }
    }

    function getOOXML() {
        // Get a reference to the Div where we will write the status of our operation.
        const report = document.getElementById("status");
        const textArea = document.getElementById("dataOOXML");
        // Remove all nodes from the status Div so we have a clean space to write to.
        while (report.hasChildNodes()) {
            report.removeChild(report.lastChild);
        }

        // Now we can begin the process.
        // First we call the getSelectedDataAsync method. The included parameter is the coercion
        // type (in our case ooxml).
        // Note that the optional parameters valueFormat and filterType are not relevant to this
        // method when used in Word, so they are excluded here.
        // When the method returns, the function that is provided as the second parameter will run.
        Office.context.document.getSelectedDataAsync("ooxml",
            function (result) {
                // Get a reference to our textArea element,
                // which is located at the end of the Div with the ID 'Content' in the WD_OpenXML_js.html page.

                if (result.status == "succeeded") {

                    // If the getSelectedDataAsync call succeeded, then
                    // result.value will return a valid chunk of OOXML, which we'll
                    // hold in the currentOOXML variable.

                    currentOOXML = result.value;

                    // Now we populate the text area in the task pane with the retrieved OOXML
                    // so that you can copy it out for editing.
                    // The first step below clears the text area and then we use a brief timeout to leave
                    // the text area blank momentarily and make it clear that the OOXML is being refreshed
                    // with the markup for the new selection.
                    // Then we report to the user that we were successful

                    setTimeout(function () {
                        textArea.value = currentOOXML;
                        report.innerText = "The getOOXML function succeeded!";
                    }, 400);

                    // Clear the success message after a 2-second delay.
                    setTimeout(function () {
                        report.innerText = "";
                    }, 2000);
                }
                else {
                    // This runs if the getSelectedDataAsync method does not return a success flag.
                    currentOOXML = "";
                    report.innerText = result.error.message;
                }
            });
    }

    function setOOXML() {
        // Get a reference to the Div where we will write the outcome of our operation.
        const report = document.getElementById("status");

        // Sets the currentOOXML variable to the current contents of the task pane text area.
        const textArea = document.getElementById("dataOOXML");
        currentOOXML = textArea.value;

        // Remove all nodes from the status Div so we have a clean space to write to.
        while (report.hasChildNodes()) {
            report.removeChild(report.lastChild);
        }

        // Check whether we have OOXML in the variable.
        if (currentOOXML != "") {

            // Call the setSelectedDataAsync, with parameters of:
            // 1. The Data to insert.
            // 2. The coercion type for that data.
            // 3. A callback function that lets us know if it succeeded.


            Office.context.document.setSelectedDataAsync(
                currentOOXML, { coercionType: "ooxml" },
                function (result) {
                    // Tell the user we succeeded and then clear the message after a 2-second delay.
                    if (result.status == "succeeded") {
                        report.innerText = "The setOOXML function succeeded!";
                        setTimeout(function () {
                            report.innerText = "";
                        }, 2000);
                    }
                    else {
                        // This runs if the setSelectedDataAsync method does not return a success flag.
                        report.innerText = result.error.message;

                        // Clear the text area just so we don't give you the impression that there's
                        // valid OOXML waiting to be inserted... 
                        textArea.value = "";
                    }
                });
        }
        else {

            // If currentOOXML == "" then we should not even try to insert it, because
            // that is guaranteed to cause an exception needlessly.
            report.innerText = "There is currently no OOXML to insert!"
                + " Please select some of your document and click [Get OOXML] first!";
        }
    }
})();