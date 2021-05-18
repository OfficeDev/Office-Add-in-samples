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
  createdDate: "5/18/2022 10:00:00 AM"
description: "This sample shows how to insert a template from an external Excel file and populate it with JSON data."
---

# Insert an external Excel file and populate it with JSON data

## Summary

This sample shows how to insert an existing template from an external Excel file into the currently open Excel file. Then it retrieves data from a JSON web service and populates the template for the customer.

> **Note:** The features used in this sample are currently in preview and subject to change. They are not currently supported for use in production environments. To try the preview features, you'll need to [join Office Insider](https://insider.office.com/join). A good way to try out preview features is to sign up for an Office 365 subscription. If you don't already have an Office 365 subscription, get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).

## Features

- Use **insertWorksheetsFromBase64** to insert a worksheet from another Excel file into the open Excel file.
- Get JSON data and add it to the worksheet.

## Applies to

- Excel on Windows, Mac, and on the web.

## Prerequisites

To use this sample, you'll need to [join Office Insider](https://insider.office.com/join).

## Solution

Solution | Author(s)
---------|----------
Insert an external Excel file and populate it with JSON data | Microsoft

## Version history

Version  | Date | Comments
---------| -----| --------
1.0 | 5-18-2021 | Initial release

----------

## Run the sample

To run the sample you just need to sideload the manifest. The add-in web files are served from this repo on GitHub.

1. Download the **manifest.xml** and **SalesTemplate.xlsx** files from this sample to a folder on your computer.
1. Open [Office on the web](https://office.live.com/).
1. Choose **Excel**, and then open a new document.
1. Open the **Insert** tab on the ribbon and choose **Office Add-ins**.
1. On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
   ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in"](../../images/office-add-ins-my-account.png)
1. Browse to the add-in manifest file, and then select **Upload**.
   ![The upload add-in dialog with buttons for browse, upload, and cancel.
](../../images/upload-add-in.png)
1. Verify that the add-in loaded successfully. You will see a **PnP Insert Excel file** button on the **Home** tab on the ribbon.

Once the add-in is loaded use the following steps to try out the functionality.

1. On the **Home** ribbon, choose **PnP Insert Excel file**.
1. In the task pane, select the **Choose file** button.
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

      // We know that the table in this template starts at B5, so we start with that.
      // Next, we calculate the total number of rows from our sales data.
      const startRow = 5;
      var address = "B" + startRow + ":F" + (newSalesData.length + startRow - 1);
      // Write the sales data to table in the template.
      var range = sheet.getRange(address);
      range.values = newSalesData;
      sheet.activate();
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