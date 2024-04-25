/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
/**
 * Basic function to show how to insert a value into cell A1 on the selected Excel worksheet.
 */
console.log("Loading Index.razor.js");

export function helloButton() {

    console.log("We are now entering function: helloButton");

    return Excel.run(context => {


        // Insert text 'Hello world!' into cell A1.
        context.workbook.worksheets.getActiveWorksheet().getRange("A1").values = [['Hello world!']];

        // sync the context to run the previous API call, and return.
        return context.sync();
    });
}
