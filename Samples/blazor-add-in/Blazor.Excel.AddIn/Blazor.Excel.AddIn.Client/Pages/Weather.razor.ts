/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
/**
 * Basic function to show how to insert a value into cell A1 on the selected Excel worksheet.
 */
export function copyButton(forecasts: any) {
  console.log("We are now entering function: copyButton");

  return Excel.run(context => {
    let sheet: Excel.Worksheet = context.workbook.worksheets.getActiveWorksheet();
    let expensesTable: Excel.Table = sheet.tables.add("A1:D1", true /*hasHeaders*/);

    expensesTable.getHeaderRowRange().values= [["Date", "Temp. (C)", "Temp. (F)", "Summary"]];

    console.log(forecasts);

    expensesTable.rows.add(undefined /*add rows to the end of the table*/, forecasts);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
  });
}