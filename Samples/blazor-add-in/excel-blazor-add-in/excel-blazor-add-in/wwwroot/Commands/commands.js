/*
 * Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 * 
 */

console.log("Loading command.js");

/* global global, Office, self, window */

/**
 * Writes the event source id to the document when ExecuteFunction runs.
 * @param event {Office.AddinCommands.Event}
 */
function writeValue(event) {
    Office.context.document.setSelectedDataAsync(
        "ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log("writeValue Failed");
            } else {
                console.log("writeValue Succeeded");
            }
        }
    );

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}

/**
 * Calls the JSInvokable function CreateBubbles to create a bubble chart, after adding the data in the active worksheet.
 * @param event {Office.AddinCommands.Event}
 */
async function createBubbles(event) {

    console.log("Running createBubbles");

    // Implement your custom code here. The following code is a simple Excel example.
    try {

        // Call JSInvokable Function here ...
        await DotNet.invokeMethodAsync("BlazorAddIn", "CreateBubbles");

        console.log("Finished createBubbles")

    } catch (error) {
        // Note: In a production add-in, notify the user through your add-in's UI.
        console.error(error);
    }

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}

/**
 * 
 * @param event {Office.AddinCommands.Event}
 */
async function highlightSelectionIndex(event) {

    // Implement your custom code here. The following code is a simple Excel example.  
    try {
        console.log("Running highlightSelectionIndex");

        console.log("Before callStaticLocalComponentMethodinit");
        await callStaticLocalComponentMethodinit("SayHelloIndex");
        console.log("After callStaticLocalComponentMethodinit");

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "red";
            await context.sync();
        });

    } catch (error) {
        // Note: In a production add-in, notify the user through your add-in's UI.
        console.error(error);
    }

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}

async function highlightSelectionBubble(event) {

    // Implement your custom code here. The following code is a simple Excel example.  
    try {
        console.log("Running highlightSelectionBubble");

        console.log("Before callStaticLocalComponentMethodinit");
        await callStaticLocalComponentMethodinit("SayHelloBubble");
        console.log("After callStaticLocalComponentMethodinit");

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "red";
            await context.sync();
        });

    } catch (error) {
        // Note: In a production add-in, notify the user through your add-in's UI.
        console.error(error);
    }

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}

async function callStaticLocalComponentMethodinit(methodname) {

    console.log("In callStaticLocalComponentMethodinit");

    try {
        let name = "Initializing";

        // Call JSInvokable Function here ...
        name = await DotNet.invokeMethodAsync("BlazorAddIn", methodname, "Blazor Fan");
        console.log("Finished Initializing: " + name)

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.values = [[name]];
            await context.sync();
        });

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "yellow";
            await context.sync();
        });
    }
    catch (err) {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.values = [[err.message]];
            await context.sync();
        });
        console.log();
        console.log("Error call : " + err.message);
    }
    finally {
        console.log("Finish callStaticLocalComponentMethodinit");
    }
}

// Associate the functions with their named counterparts in the manifest XML.
Office.actions.associate("writeValue", writeValue);
Office.actions.associate("createBubbles", createBubbles);
Office.actions.associate("highlightSelectionIndex", highlightSelectionIndex);
Office.actions.associate("highlightSelectionBubble", highlightSelectionBubble);