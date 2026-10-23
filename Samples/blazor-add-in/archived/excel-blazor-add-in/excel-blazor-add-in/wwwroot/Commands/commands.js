/*
 * Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 * 
 */

console.log("Loading command.js");

/**
 * Writes the event source id to the document when ExecuteFunction runs.
 * @param event {Office.AddinCommands.Event}
 */
async function writeValue(event) {

    console.log("In writeValue");

    try {
        let message = "ExecuteFunction works. Button ID=" + event.source.id;

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.values = [[message]];
            range.getEntireColumn().format.autofitColumns();
            await context.sync();
        });

        console.log("writeValue Succeeded");

    } catch (err) {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            const cellRange = range.getCell(0, 0);
            cellRange.values = [[err.message]];
            await context.sync();
        });
        console.log();
        console.log("Error call : " + err.message);
    }
    finally {
        console.log("Finish callStaticLocalComponentMethodinit");
    }

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
 * Writes the text from the Index Blazor Page to the Worksheet when highlightSelectionIndex runs.
 * @param event {Office.AddinCommands.Event}
 */
async function highlightSelectionIndex(event) {

    // Implement your custom code here. The following code is a simple Excel example.  
    try {
        console.log("Running highlightSelectionIndex");

        console.log("Before callStaticLocalComponentMethodinit");
        await callStaticLocalComponentMethodinit("SayHelloIndex");
        console.log("After callStaticLocalComponentMethodinit");

        // Used to verify the previous function call, if that fails, this will not run.
        // It will be skipped on error and jump into the catch block.
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "LightBlue";
            await context.sync();
        });

    } catch (error) {
        // Note: In a production add-in, notify the user through your add-in's UI.
        console.error(error);
    }

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}

/**
 * Writes the text from the BubbleChart Blazor Page to the Worksheet when highlightSelectionIndex runs.
 * @param event {Office.AddinCommands.Event}
 */
async function highlightSelectionBubble(event) {

    // Implement your custom code here. The following code is a simple Excel example.  
    try {
        console.log("Running highlightSelectionBubble");

        console.log("Before callStaticLocalComponentMethodinit");
        await callStaticLocalComponentMethodinit("SayHelloBubble");
        console.log("After callStaticLocalComponentMethodinit");

        // Used to verify the previous function call, if that fails, this will not run.
        // It will be skipped on error and jump into the catch block.
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "LightBlue";
            await context.sync();
        });

    } catch (error) {

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "Red";
            await context.sync();
        });

        // Note: In a production add-in, notify the user through your add-in's UI.
        console.error(error);
    }

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}

/**
 * Local function to call the JSInvokable function in the Blazor Component.
 * @param methodname {methodname}
 */
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
            range.getEntireColumn().format.autofitColumns();
            await context.sync();
        });

        // Used to verify the previous function call, if that fails, this will not run.
        // It will be skipped on error and jump into the catch block.
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "yellow";
            await context.sync();
        });
    }
    catch (err) {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            const cellRange = range.getCell(0, 0);
            cellRange.values = [[err.message]];
            await context.sync();
        });

        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = "red";
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