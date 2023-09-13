// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const dataSourceUrl =
  "https://officedev.github.io/Office-Add-in-samples/Samples/excel-insert-file";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    let fileInput = document.getElementById("fileInput");
    fileInput.addEventListener("change", insertSheets);
  }
});

async function insertSheets() {
  const myFile = document.getElementById("fileInput");
  const reader = new FileReader();

  reader.onload = async (event) => {
    Excel.run(async (context) => {
      try {
        // Remove the metadata before the base64-encoded string.
        const startIndex = reader.result.toString().indexOf("base64,");

        // 7 is the length of the "base64," string to skip past
        const workbookContents = reader.result
          .toString()
          .substr(startIndex + 7);

        // STEP 1: Insert the template into the workbook.
        const workbook = context.workbook;

        // Set up the insert options.
        const options = {
          sheetNamesToInsert: ["Template"], // Insert the "Template" worksheet from the source workbook.
          positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
          relativeTo: "Sheet1",
        }; // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.

        // Insert the external worksheet.
        workbook.insertWorksheetsFromBase64(workbookContents, options);

        // In Excel on the web, if the worksheet being inserted contains unsupported features,
        // such as Comment, Slicer, Chart, and PivotTable, insertWorksheetsFromBase64 will fail.
        // In your production add-in, you should notify the user in the add-ins UI.
        // As a workaround they can use Excel on desktop, or choose a different worksheet.
        await context.sync();

        // STEP 2: Add data from the "Service".
        const sheet = context.workbook.worksheets.getItem("Template");

        // Get data from your REST API. For this sample, the JSON is fetched from a file in the repo.
        let json;
        let response = await fetch(dataSourceUrl + "/data.json");
        if (response.ok) {
          json = await response.json();
        } else {
          console.error("HTTP-Error: " + response.status);
          return;
        }

        // Map JSON to table columns.
        const newSalesData = json.salesData.map((item) => [
          item.PRODUCT,
          item.QTR1,
          item.QTR2,
          item.QTR3,
          item.QTR4,
        ]);

        // We know that the table in this template starts at B5, so we start with that.
        // Next, we calculate the total number of rows from our sales data.
        const startRow = 5;
        const address =
          "B" + startRow + ":F" + (newSalesData.length + startRow - 1);

        // Write the sales data to the table in the template.
        const range = sheet.getRange(address);
        range.values = newSalesData;
        sheet.activate();
        return context.sync();
      } catch (error) {
        // In your production add-in, you should notify the user in the add-in UI.
        console.error(error);
        return;
      }
    });
  };

  // Read the file as a data URL so that we can parse the base64-encoded string.
  reader.readAsDataURL(myFile.files[0]);
}
