---
page_type: sample
products:
- office-excel
- office
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: "5/14/2022 10:00:00 AM"
description: "This sample shows how to insert a template from an external Excel file and populate it with JSON data."
---

# Insert an external Excel file and populate it with JSON data

## Summary

This sample shows how to insert an existing template from an external Excel file into the currently open Excel file. Then it retrieves data from a JSON web service and populates the template for the customer.

## Features

- Use **insertWorksheetsFromBase64** to insert a worksheet from another Excel file into the open Excel file.
- Get JSON data and add it to the worksheet.

## Applies to

- Excel on Windows, and Mac.

## Prerequisites

- Microsoft 365

## Solution

Solution | Author(s)
---------|----------
Insert an external Excel file and populate it with JSON data | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 5-14-2021 | Initial release

----------

## Run the sample

You can run this sample in Excel on Windows, or Mac. The add-in web files are served from this repo on GitHub.

1. Download the **manifest.xml** and **SalesTemplate.xlsx** files from this sample to a folder on your computer.
1. If you are using Excel on Windows, create a network share and sideload the manifest by following the instructions in [Sideload Office Add-ins for testing from a network share](https://docs.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).
1. If you are using Excel on Mac, sideload the manifest by following the instructions in [Sideload Office Add-ins on iPad and Mac for testing](https://docs.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad-and-mac#sideload-an-add-in-in-office-on-mac). Note that the instructions are for Word, but they mirror the same steps as for Excel.

Once the add-in is loaded use the following steps to try out the functionality.

1. On the **Home** ribbon, choose **Show task pane**.
2. In the task pane, select the **Choose file** button.
1. In the dialog box that opens, select the **SalesTemplate.xlsx** file that you downloaded previously. The choose **Open**.

A **Contoso Sales Report** will be inserted with a table and chart populated with data.

## Key parts of this sample

When you select the **SalesTemplate.xlsx** file, the following code in **index.js** inserts the template. It sets up an object named **options** to identify the sheet by name (**Template**). Then it calls the Office.js **insertWorksheetsFromBase64** API to insert the template into the current worksheet.

```javascript
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
```

Next, it gets the JSON which is in the **data.json** file in this repo.

```javascript
 // STEP 2: Add data from the "Service".
      const sheet = context.workbook.worksheets.getItem("Template");

      // Get data from your REST API. For this sample, the JSON is fetched from a file in the repo.
      let response = await fetch(dataSourceUrl + "/data.json");
      if (response.ok) {
        var json = await response.json();
      } else {
        console.error("HTTP-Error: " + response.status);
      }

```

Finally it adds the JSON to the table.

```javascript
 //map JSON to table columns
      const newSalesData = json.salesData.map((item) => [
        item.PRODUCT,
        item.QTR1,
        item.QTR2,
        item.QTR3,
        item.QTR4,
        "",
      ]);

      const salesTable = sheet.tables.getItem("SalesTable");
      salesTable.rows.add(null, newSalesData);
```

## Run the sample from Localhost

If you prefer to host the web server for the sample on your computer, follow these steps:

1. Open the **index.js** file.
1. Edit line 4 to refer to the localhost:3000 endpoint as shown in the following code.
    
    ```javascript
    const dataSourceUrl = "https://localhost:3000";
    ```
    
1. Save the file.
1. You need http-server to run the local web server. If you haven't installed this yet you can do this with the following command:
    
    ```console
    npm install --global http-server
    ```
    
2. Use a tool such as openssl to generate a self-signed certificate that you can use for the web server. Move the cert.pem and key.pem files to the webworker-customfunction folder for this sample.
3. From a command prompt, go to the web-worker folder and run the following command:
    
    ```console
    http-server -S --cors . -p 3000
    ```
    
4. To reroute to localhost run office-addin-https-reverse-proxy. If you haven't installed this you can do this with the following command:
    
    ```console
    npm install --global office-addin-https-reverse-proxy
    ```
    
    To reroute run the following in another command prompt:
    
    ```console
    office-addin-https-reverse-proxy --url http://localhost:3000
    ```
    
5. Follow the steps in [Run the sample](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts#run-the-sample), but upload the `manifest-localhost.xml` file for step 6.

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://telemetry.sharepointpnp.com/pnp-officeaddins/samples/excel-insert-external-file" />