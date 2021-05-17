// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const dataSourceUrl = "https://davidchesnut.github.io";

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
      // Remove the metadata before the base64-encoded string.
      const startIndex = reader.result.toString().indexOf("base64,");

      // 7 is the length of the "base64," string to skip past
      const workbookContents = reader.result.toString().substr(startIndex + 7);

      // STEP 1: Insert the template into the workbook.
      const workbook = context.workbook;

      // Set up the insert options.
      var options = {
        sheetNamesToInsert: ["Template"], // Insert the "Template" worksheet from the source workbook.
        positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
        relativeTo: "Sheet1",
      }; // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.

      // Insert the external worksheet.
      workbook.insertWorksheetsFromBase64(workbookContents, options);
      await context.sync();

      // STEP 2: Add data from the "Service".
      const sheet = context.workbook.worksheets.getItem("Template");

      // Get data from your REST API. For this sample, the JSON is fetched from a file in the repo.
      let response = await fetch(dataSourceUrl + "/data.json");
      if (response.ok) {
        var json = await response.json();
      } else {
        console.error("HTTP-Error: " + response.status);
      }

      //map JSON to table columns
      const newSalesData = json.salesData.map((item) => [
        item.PRODUCT,
        item.QTR1,
        item.QTR2,
        item.QTR3,
        item.QTR4,
        "",
      ]);

      //insert data as new rows in table.
      const salesTable = sheet.tables.getItem("SalesTable");
      salesTable.rows.add(null, newSalesData);

      return context.sync();
    });
  };

  // Read the file as a data URL so that we can parse the base64-encoded string.
  reader.readAsDataURL(myFile.files[0]);
}
