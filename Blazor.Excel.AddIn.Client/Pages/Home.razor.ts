/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

console.log("Loading Home.razor.ts");

/**
 * Basic function to show how to insert a value into cell A1 on the selected Excel worksheet.
 * Inserts 'Hello world!!!' text into cell A1 of the active worksheet.
 */

export async function insertText() {

  console.log("We are now entering function: insertText");

  try {
    await Excel.run(async function (context) {

      // Insert text 'Hello world!' into cell A1.
      const activeWorksheet: Excel.Worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range: Excel.Range = activeWorksheet.getRange("A1");
      range.values = [['Hello world!!!']];

      console.log("Welcome text created successfully.");

      // sync the context to run the previous API call, and return.
      await context.sync();
    });
  } catch (error: unknown) {
    const errorMessage: string = error instanceof Error ? error.message : String(error);
    console.error("Error creating welcome: ", errorMessage);
  }
}