Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = insertSheets;
  }
});

async function insertSheets() {
  const myFile = document.getElementById("file");
  const reader = new FileReader();

  reader.onload = (event) => {
    Excel.run((context) => {
      // Remove the metadata before the base64-encoded string.
      const startIndex = reader.result.toString().indexOf("base64,");
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

      // Get data from your REST API.
      // For simplicity, the JSON data is hardcoded in this sample and retrieved from an object called 'salesData'.

      let response = await fetch("url");

      if (response.ok) {
        // if HTTP-status is 200-299
        // get the response body (the method explained below)
        let json = await response.json();
      } else {
        console.error("HTTP-Error: " + response.status);
      }
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
      var address =
        "B" + startRow + ":F" + (newSalesData.length + startRow - 1);

      // Write the sales data to table in the template.
      var range = sheet.getRange(address);
      range.values = newSalesData;

      return context.sync().catch(function (error) {
        console.log(error);
      });
    });
  };

  // Read the file as a data URL so that we can parse the base64-encoded string.
  reader.readAsDataURL(myFile.files[0]);
}
